param([parameter(Position=0, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true, mandatory=$true)][string]$Identity, 
      [parameter(Position=1, mandatory=$true)][string]$RemoteForestDomainController, 
      [parameter(Position=2, mandatory=$true)][Management.Automation.PSCredential]$RemoteForestCredential, 
      [string]$LocalForestDomainController,
      [Management.Automation.PSCredential]$LocalForestCredential,
      [string]$TargetMailUserOU,
      [string]$MailboxDeliveryDomain,
      [switch]$LinkedMailUser,
      [switch]$DisableEmailAddressPolicy,
      [switch]$UseLocalObject,
      [switch]$OverwriteLocalObject)

begin
{
    # ---------------------------------------------------------------------------------------------------
    function findADObject($searchRoot, $filter)
    # ---------------------------------------------------------------------------------------------------
    {
        $searcher = new-object System.DirectoryServices.DirectorySearcher($searchRoot)
        $searcher.filter = $filter
        $user = $searcher.findall()

        if ($user -eq $null -or $user.count -eq 0)
        {
            return $null
        }
        elseif ($user.count -gt 1)
        {
            foreach ($usr in $user)
            {
                Write-Warning ("Object Found:" + $usr.GetDirectoryEntry().distinguishedName)
            }
            Write-Host "Tips: For Source object, please check the duplication of distinguishedName, mailNickName, displayName, objectGUID and proxyAddresses. You could parse the objectGUID in the Identity parameter to make it unique instead."
            Write-Host "Tips: For Local object, please check whether the proxyAddresses have duplicated item in these objects. You need to correct the dirty data before you could continue."
            throw "Multiple objects found in AD."
        }
        else
        {
            return $user[0].GetDirectoryEntry()
        }
    }

    # ---------------------------------------------------------------------------------------------------
    function checkUserExist ($OU, $filter)
    # ---------------------------------------------------------------------------------------------------
    {
        $searcher = new-object System.DirectoryServices.DirectorySearcher($OU)
        $searcher.filter = $filter
        $user = $searcher.findone() 
        if ($user -eq $null -or $user.count -eq 0)
        {
            return $false
        }
        else
        {
            return $true
        }
    }

    # ---------------------------------------------------------------------------------------------------
    function getAttributeFriendlyValue ($attribute, $attValue)
    # ---------------------------------------------------------------------------------------------------
    {
        if (($attribute -eq "msExchMailboxGuid") -or ($attribute -eq "msExchArchiveGuid"))
        {
            return ([System.Guid]($attValue)).ToString()
        }
        else
        {
            return $attValue
        }
    }

    # ---------------------------------------------------------------------------------------------------
    function copyIfExist ($target, [array]$attriblist, $propertybag)
    # ---------------------------------------------------------------------------------------------------
    {
        foreach($att in $attriblist)
        {
            if ($propertybag.Contains($att))
            {
                $friendlyValue = getAttributeFriendlyValue $att $propertybag.Item($att).Value
                Write-Verbose "Setting $att to $friendlyValue"
                [void]($target.Put($att, $propertybag.Item($att).Value))
            }
        }
    }

    # ---------------------------------------------------------------------------------------------------
    function getEscapedldapFilterStr ([string]$original)
    # ---------------------------------------------------------------------------------------------------
    {
        $escape = $original.replace("\", "\5c")
        $escape = $escape.replace("(", "\28").replace(")", "\29")
        $escape = $escape.replace("&", "\26").replace("|", "\7c")
        $escape = $escape.replace("=", "\3d").replace(">", "\3e")
        $escape = $escape.replace("<", "\3c").replace("~", "\7e")
        $escape = $escape.replace("*", "\2a").replace("/", "\2f")
        return $escape
    }

    # ---------------------------------------------------------------------------------------------------
    function getEscapedDNStr ([string]$original)
    # ---------------------------------------------------------------------------------------------------
    {
        $escape = $original.replace(",", "\,")
        $escape = $escape.replace("+", "\+")
        $escape = $escape.replace("""", "\""")
        $escape = $escape.replace("#", "\#")
        $escape = $escape.replace(";", "\;")
        return $escape
    }

    # ---------------------------------------------------------------------------------------------------
    function sidToLDAPQuery([byte[]]$sid)
    # ---------------------------------------------------------------------------------------------------
    {
        foreach ($by in $sid)
        {
            $ret += "\" + $by.tostring("X")
        }
        return $ret
    }
    
    # ---------------------------------------------------------------------------------------------------    
    function MasterAccountSidIsSelf ( $srcMbxAttributes )
    # ---------------------------------------------------------------------------------------------------
    {
        if ($srcMbxAttributes.Contains("msExchMasterAccountSid"))
        {
            $master = new-object System.Security.Principal.SecurityIdentifier($srcMbxAttributes.Item("msExchMasterAccountSid").value, 0)
            if ($master.IsWellKnown("SelfSid"))
            {
                return $true
            }
        }
        return $false
    }

    # ---------------------------------------------------------------------------------------------------
    function findLocalObject ($OU, $srcuser)
    # ---------------------------------------------------------------------------------------------------
    {
        $objectClassFilter = "(&(!objectClass=computer)(|(objectClass=user)(objectClass=contact)(objectClass=group)(objectClass=msExchDynamicDistributionList)))"
        $usr = $null
        if ($srcuser.properties.Contains("msExchMasterAccountSid") -and -not (MasterAccountSidIsSelf $srcuser.properties))
        {
            $sourcesid += sidToLDAPQuery $srcuser.properties.Item("msExchMasterAccountSid").Value
            $filter = "(| (ObjectSid=$sourcesid) (msExchMasterAccountSid=$sourcesid) )"
            $filter = "(&($objectClassFilter)($filter))"

            $usr = findADObject $OU $filter
        }
        if ($usr -eq $null)
        {
            $address = $srcuser.Properties.Item("proxyAddresses")
            foreach ($addr in $address)
            {
                if ($addr.startswith("x500:", "OrdinalIgnoreCase") -or $addr.startswith("smtp:", "OrdinalIgnoreCase"))
                {
                    $addr1 = getEscapedldapFilterStr ($addr.Substring(0,4).toUpper() + $addr.Substring(4))
                    $addr2 = getEscapedldapFilterStr ($addr.Substring(0,4).toLower() + $addr.Substring(4))
                    $filterstring += "(proxyAddresses=$addr1) (proxyAddresses=$addr2)"
                }
            }
            
            $filter = "(| $filterstring)"
            $filter = "(&($objectClassFilter)($filter))"

            $usr = findADObject $OU $filter
        }
        return $usr
    }

    # ---------------------------------------------------------------------------------------------------
    function generateUniqueSAM ($ou, $srcMbxAttributes)
    # ---------------------------------------------------------------------------------------------------
    {
        $uniquesam = $srcMbxAttributes.Item('samaccountname').Value
        $retrycount = 30
        if ($uniquesam.Length -lt 20)
        {
            while ($retrycount -gt 0 -and (checkUserExist $ou "(samAccountName=$(getEscapedldapFilterStr $uniquesam))"))
            {
                $uniquesam = $srcMbxAttributes.Item("samaccountname").Value + (random)
                if ($uniquesam.length -gt 20)
                {
                    $uniquesam = $uniquesam.substring(0,20)
                }
                $retrycount = $retrycount - 1
            }
        }
        return $uniquesam
    }

    # ---------------------------------------------------------------------------------------------------
    function generateUniqueUPN ($ou, $srcMbxAttributes, $fallbacks)
    # ---------------------------------------------------------------------------------------------------
    {
        if ($srcMbxAttributes.Contains('userPrincipalName'))
        {
            $uniqueupn = $srcMbxAttributes.Item('userPrincipalName').Value
            if ($uniqueupn -match "^(.*)(@.*)$")
            {
                $postfix = $matches[2]
                $prefix  = $matches[1]
            }
            $preferedupn = ,$uniqueupn + $fallbacks
            foreach ($upn in $preferedupn)
            {
                if ($upn -ne $null)
                {
                    if ($upn.contains("@"))
                    {
                        $testupn = $upn
                    }
                    else
                    {
                        $testupn = "$upn$postfix"
                    }
                    if ($(checkUserExist $ou "(userPrincipalName=$(getEscapedldapFilterStr $testupn))") -eq $false)
                    {
                        return $testupn
                    }
                }
            }
            #try to use prefered upn, if all unsuitable, generate a new one
            while ($(checkUserExist $ou "(userPrincipalName=$(getEscapedldapFilterStr $uniqueupn))"))
            {
                $uniqueupn = "$prefix$(random)$postfix"
            }
        }
        return $uniqueupn
    }

    # ---------------------------------------------------------------------------------------------------
    function copyBasicAttributes ($newuser, $srcAttributes)
    # ---------------------------------------------------------------------------------------------------
    {
        $copyAttributes="displayName",
                        "Mail",
                        "mailNickName",
                        "msExchMailboxGuid",
                        "msExchArchiveGuid",
                        "msExchUserCulture",
                        "msExchArchivename"

        [void](copyIfExist $newuser $copyAttributes $srcAttributes)
    }


    # ---------------------------------------------------------------------------------------------------
    function copyMandatoryAttributes ($newuser, $srcAttributes, $localDC)
    # ---------------------------------------------------------------------------------------------------
    {
        copyBasicAttributes $newuser $srcAttributes

        # Handle proxyAddresses specially (only copied when the local object is created at first).
        [void](copyIfExist $newuser "proxyAddresses" $srcAttributes)

        $specialAttributes = @{  "msExchRecipientDisplayType"=0x80000006;
                                 "msExchRecipientTypeDetails"=0x80;
                                 "msExchVersion"="44220983382016";
                                 "userAccountControl"=0x202 #ACCOUNTDISABLE | NORMAL_ACCOUNT
                               }

        if ($localDC -ne $null)
        {
            $specialAttributes["samaccountname"] = generateUniqueSAM $localDC $srcAttributes
            $specialAttributes["userPrincipalName"] = generateUniqueUPN $localDC $srcAttributes $newuser.cn,$specialAttributes["samaccountname"]
        }

        foreach($att in $specialAttributes.getenumerator())
        {
            if ($att.value -ne $null)
            {
                Write-Verbose "Setting $($att.key) to $($att.value)"
                [void]($newuser.put($att.key, $att.value.tostring()))
            }
        }
    }

    # ---------------------------------------------------------------------------------------------------
    function getTargetAddress ($srcProxyAddresses, $writeErrorIfNotExist)
    # ---------------------------------------------------------------------------------------------------
    {
        foreach ($addr in $srcProxyAddresses)
        {
            if ($addr -match "^(SMTP|smtp):.*@(.*)$")
            {
                #if don't specify authoritative domains, use primary smtp address
                if (([string]::IsNullOrEmpty($MailboxDeliveryDomain) -and $addr.startswith("SMTP")) -or
                    ($matches[2] -eq $MailboxDeliveryDomain))
                {
                    Write-Verbose "Setting targetAddress to $addr"
                    return $addr
                }
            }
        }

        if ($writeErrorIfNotExist)
        {
            $errFixTip = $null
            if ([string]::IsNullOrEmpty($MailboxDeliveryDomain))
            {
                $errFixTip = "PrimaryEmailAddress exists"
            }
            else
            {
                $errFixTip = "some EmailAddress matches MailboxDeliveryDomain $MailboxDeliveryDomain"
            }

            # Terminate if fail to determine targetAddress (necessary for MEU)
            throw "Unable to determine the targetAddress for the newly created MEU. Please ensure that $errFixTip in source object."
        }
    }

    # ---------------------------------------------------------------------------------------------------
    function disableEmailAddressPolicy ($user)
    # ---------------------------------------------------------------------------------------------------
    {
        if ($DisableEmailAddressPolicy)
        {
            Write-Verbose "Disable EmailAddressPolicy."
            try
            {
                $dcParameter = @{};
                if ($LocalForestDomainController -ne $null)
                {
                    $dcParameter = @{DomainController=$LocalForestDomainController}
                }
                Set-MailUser $user.distinguishedName.Value @dcParameter -EmailAddressPolicyEnabled:$false
            }
            catch
            {
                 # Terminate if fail to disable the EAP.
                 throw "Fail to disable the EmailAddressPolicy for MEU($($user.distinguishedName)). Error: $($Error[0])"
            }
        }
    }

    # ---------------------------------------------------------------------------------------------------
    function generateLegacyExchangeDN ($user)
    # ---------------------------------------------------------------------------------------------------
    {
        # Update the LegacyExchangeDN if it doesn't exist.
        if ($user.properties.Contains("LegacyExchangeDN") -eq $false)
        {
            Write-Verbose "Invoke Update-Recipient to Update LegacyExchangeDN."
            try
            {
                Update-Recipient $user.distinguishedName.Value @DomainControllerParameterSet
                #rebind ad object to retrieve new properties set by Update-Recipient
                $user.RefreshCache([array]"legacyExchangeDN")
            }
            catch
            {
                Write-Error "Error updating recipient MEU($($user.DistinguishedName)) to generate the legacyDN. Please fix the error, run the Update-Recipient task to generate the LegacyDN and this script again. Error: $($Error[0])"
                return
            }
        }
    }

    # ---------------------------------------------------------------------------------------------------
    function mergeLegacyDNToProxyAddress ($user, $X500ProxyAddresses, $updateImmediately, $sourceOrLocal)
    # ---------------------------------------------------------------------------------------------------
    {
        # Merge legacyDN to source object
        if ($X500ProxyAddresses.Length -gt 0)
        {
            $userFriendlyName = "Object($($user.distinguishedName))"
            if ([string]::IsNullOrEmpty($user.distinguishedName))
            {
                $userFriendlyName = "New Object" # DN is not available.
            }
            $updateRequired = $false
            foreach ($X500ProxyAddress in $X500ProxyAddresses)
            {
                if ($user.Properties.Item("proxyAddresses").tostring().toupper().Contains($X500ProxyAddress.ToUpper()) -eq $false)
                {
                    Write-Host "Appending $X500ProxyAddress to proxyAddresses of $userFriendlyName in $sourceOrLocal forest." -ForegroundColor Green
                    [void]($user.putex(3, "proxyAddresses", [array]$X500ProxyAddress))
                    $updateRequired = $true
                }
            }
            if ($updateRequired -and $updateImmediately)
            {
                try
                {
                    $user.setinfo() # Might get Access Denied
                }
                catch
                {
                    Write-Error "Error appending $X500ProxyAddress to proxyAddresses of $userFriendlyName in $sourceOrLocal forest. Please update the proxyAddresses manually, or fix the error and run this script again. Error: $($Error[0])"
                    return
                }
            }
        }
    }

    # ---------------------------------------------------------------------------------------------------
    function generateNewCN ($localDC, $oldcn, $oudn)
    # ---------------------------------------------------------------------------------------------------
    {
            $newcn = getEscapedldapFilterStr $oldcn
            $newcn = getEscapedDNStr $newcn
            $tryCnt = 0;
            $oldcnLength = $oldcn.Length;
            $localPath = $localdc.Path;
            # Usage of this script may not provide the LocalForestDomainController parameter, in which case, the $localdc.Path is empty.
            if ([string]::IsNullorEmpty($localPath))
            {
                $localPath = "LDAP:/";
            }
            while ([DirectoryServices.DirectoryEntry]::exists("$localPath/cn=$newcn,$oudn"))
            {
                if ($tryCnt -gt 10)
                {
                    throw "Unable to generate new CN, the script has tried 10 times with random string appended. You could re-run again if you're sure it could be generated.";
                }
                $tryCnt = $tryCnt + 1;

                if ($oldcnLength -ge 64)
                {
                    # if old CN legth is already maximum (64), we can't generate new one.
                    throw "Unable to generate new CN. The original CN $newcn exists and its length >= 64, thus we can't append random string at the end to generate new one.";
                }

                $ranStr = (random).ToString()
                if ($oldcnLength + $ranStr.Length -gt 64)
                {
                    $ranStr = $ranStr.subString(0, 64 - $oldcnLength);
                }
                $newcn = getEscapedldapFilterStr ($oldcn + $ranStr)
                $newcn = getEscapedDNStr $newcn
            }
            return $newcn;
    }

    # ---------------------------------------------------------------------------------------------------
    function createMailUserAccount ($localDC, $ou, $srcMbxAttributes)
    # ---------------------------------------------------------------------------------------------------
    {
        try{
            $newcn = generateNewCN $localDC $srcMbxAttributes.Item("cn").value $ou.distinguishedname

            [void]($newuser = $ou.create("user", "cn=$newcn"))
            
            copyMandatoryAttributes $newuser $srcMbxAttributes $localDC

            #additional operations for proxyaddresses and targetaddress

            if ($srcMbxAttributes.Contains("LegacyExchangeDN"))
            {
                $X500proxyAddr = "x500:" + $srcMbxAttributes.Item("LegacyExchangeDN").value
                mergeLegacyDNToProxyAddress $newuser ([array]$X500proxyAddr) $false "Local"
            }

            $srcProxyAddresses = $srcMbxAttributes.Item("proxyAddresses")
            $targetAddress = getTargetAddress $srcProxyAddresses $true
            if ($targetAddress -ne $null)
            {
                [void]($newuser.put("targetAddress", $targetAddress))
            }

            [void]($newuser.SetInfo())

            $newuser.RefreshCache([array]"distinguishedName")

            return $newuser
        }
        catch
        {
            # Terminate the script if fail to create the new user.
            throw "Error creating mailuser CN=$newcn,$($ou.distinguishedname) in local forest or setting its mandatory attributes. Error: $($Error[0])"
        }
    }

    # ---------------------------------------------------------------------------------------------------
    function copyGalySyncAttributes ($user, $srcMbxAttributes)
    # ---------------------------------------------------------------------------------------------------
    {
        $copyAttributes= "C",
                         "Co",
                         "countryCode",
                         "Company",
                         "Department",
                         "facsimileTelephoneNumber",
                         "givenName",
                         "homePhone",
                         "Info",
                         "Initials",
                         "L",
                         "Mobile",
                         "msExchAssistantName",
                         "msExchHideFromAddressLists",
                         "otherHomePhone",
                         "otherTelephone",
                         "Pager",
                         "physicalDeliveryOfficeName",
                         "postalCode",
                         "Sn",
                         "St",
                         "streetAddress",
                         "telephoneAssistant",
                         "telephoneNumber",
                         "Title"

        copyIfExist $user $copyAttributes $srcMbxAttributes
    }

    # ---------------------------------------------------------------------------------------------------
    function copyE2k7OptionalAttributes ($user, $srcMbxAttributes)
    # ---------------------------------------------------------------------------------------------------
    {
        $copyAttributes= #"Cn",
                         "Comment",
                         "deletedItemFlags",
                         "delivContLength",
                         "departmentNumber",
                         "Description",
                         "Division",
                         "employeeID",
                         "employeeNumber",
                         "employeeType",
                         "homePostalAddress",
                         "internationalISDNNumber",
                         "ipPhone",
                         "Language",
                         "localeID",
                         "mAPIRecipient",
                         "middleName",
                         "msDS-PhoneticCompanyName",
                         "msDS-PhoneticDepartment",
                         "msDS-PhoneticDisplayName",
                         "msDS-PhoneticFirstName",
                         "msDS-PhoneticLastName",
                         "msExchBlockedSendersHash",
                         "msExchELCExpirySuspensionEnd",
                         "msExchELCExpirySuspensionStart",
                         "msExchELCMailboxFlags",
                         "msExchExternalOOFOptions",
                         "msExchMessageHygieneFlags",
                         "msExchMessageHygieneSCLDeleteThreshold",
                         "msExchMessageHygieneSCLJunkThreshold",
                         "msExchMessageHygieneSCLQuarantineThreshold",
                         "msExchMessageHygieneSCLRejectThreshold",
                         "msExchMDBRulesQuota",
                         "msExchPoliciesExcluded",
                         "msExchSafeRecipientsHash",
                         "msExchSafeSendersHash",
                         "msExchUMSpokenName",
                         "O",
                         "otherFacsimileTelephoneNumber",
                         "otherIpPhone",
                         "otherMobile",
                         "otherPager",
                         "preferredDeliveryMethod",
                         "personalPager",
                         "personalTitle",
                         "Photo",
                         "pOPCharacterSet",
                         "pOPContentFormat",
                         "postalAddress",
                         "postOfficeBox",
                         "primaryInternationalISDNNumber",
                         "primaryTelexNumber",
                         "showInAdvancedViewOnly",
                         "Street",
                         "terminalServer",
                         "textEncodedORAddress",
                         "thumbnailLogo",
                         "thumbnailPhoto",
                         "url",
                         "userCert",
                         "userCertificate",
                         "userSMIMECertificate",
                         "wWWHomePage"
        foreach ($i in 1..15)
        {
            $copyAttributes += "extensionAttribute$i";
        }


        copyIfExist $user $copyAttributes $srcMbxAttributes
    }

    # ---------------------------------------------------------------------------------------------------
    function findCorrespondingADObject ($targetOU, $DN, $srcDomain)
    # ---------------------------------------------------------------------------------------------------
    {
        $cn = "$DN".substring(0, "$DN".indexof(",DC="))
        $srcreferenceobject = $srcDomain.children.find($cn)
        $usr = $null
        if ($srcreferenceobject -ne $null)
        {
            if ($srcreferenceobject.Properties.Contains("legacyExchangeDN"))
            {
                $legexch = getEscapedldapFilterStr $srcreferenceobject.Properties.Item("legacyExchangeDN")
                $addrfilter = "(proxyAddresses=x500:$legexch) (proxyAddresses=X500:$legexch)"
            }
            $address = $srcreferenceobject.Properties.Item("proxyAddresses")
            foreach ($addr in $address)
            {
                if ($addr.startswith("x500:", "OrdinalIgnoreCase"))
                {
                    $addrfilter += "(legacyExchangeDN=$(getEscapedldapFilterStr $addr.substring(5)))"
                }
                if ($addr.startswith("smtp:", "OrdinalIgnoreCase") -or $addr.startswith("x500:", "OrdinalIgnoreCase"))
                {
                    $addr1 = getEscapedldapFilterStr ($addr.Substring(0,4).toUpper() + $addr.Substring(4))
                    $addr2 = getEscapedldapFilterStr ($addr.Substring(0,4).toLower() + $addr.Substring(4))
                    $addrfilter += "(proxyAddresses=$addr1) (proxyAddresses=$addr2)"
                }
            }
            if ([string]::IsNullOrEmpty($addrfilter) -eq $false)
            {
                $filter = "(| $addrfilter)"
                $objectClassFilter = "(&(!objectClass=computer)(|(objectClass=user)(objectClass=contact)(objectClass=group)(objectClass=msExchDynamicDistributionList)))"

                $filterWithObjectClass = "(&($objectClassFilter)($filter))"
                $usr = findADObject $targetOU $filterWithObjectClass

                if ($usr -eq $null)
                {
                    #user not found, try find with loose condition. Because link/backlink may be in other object type.
                    $usr = findADObject $targetOU $filter
                }
            }

            return $usr
        }
    }

    # ---------------------------------------------------------------------------------------------------
    function setLinkedAttribute ($attribname, $backlinkname, $targetOU, $user, $srcMbxAttributes, $srcDomain)
    # ---------------------------------------------------------------------------------------------------
    {
        if ($srcMbxAttributes.contains($attribname))
        {
            foreach ($dn in $srcMbxAttributes.item($attribname))
            {
                try
                {
                    $corobj = findCorrespondingADObject $targetOU $dn $srcDomain
                    if ($corobj -eq $null)
                    {
                        Write-Warning "Cannot find corresponding object for $dn in current forest. `'$attribname`' not set."
                    }
                    else
                    {
                        Write-Verbose "Setting $attribname to $($corobj.properties.item('distinguishedname'))"
                        if (($attribname -eq "Manager") -or ($attribname -eq "managedBy") -or ($attribname -eq "altRecipient"))
                        {
                            $user.put($attribname, $($corobj.properties.item('distinguishedname')))
                        }
                        else
                        {
                            $user.putex(3, $attribname, [array]"$($corobj.properties.item('distinguishedname'))")
                        }
                    }
                }
                catch
                {
                    Write-Warning "Error updating $($user.distinguishedName)   Attribute: $attribname! Attribute Not Set! Error: $($Error[0])"
                }
            }
        }
        
        #find backlink from source MBX, set it on corresponding user in target
        if ($srcMbxAttributes.contains($backlinkname))
        {
            foreach ($dn in $srcMbxAttributes.item($backlinkname))
            {
                try
                {
                    $corobj = findCorrespondingADObject $targetOU $dn $srcDomain
                    if ($corobj -eq $null)
                    {
                        Write-Warning "Cannot find corresponding object for $dn in current forest. `'$attribname`' not updated."
                    }
                    else
                    {
                        if (($attribname -eq "Manager") -or ($attribname -eq "managedBy") -or ($attribname -eq "altRecipient"))
                        {
                            $corobj.Put($attribname, $($user.properties.item("distinguishedname")))
                        }
                        else
                        {
                            $corobj.PutEx(3, $attribname, [array]"$($user.properties.item("distinguishedname"))")
                        }
                        $corobj.SetInfo()
                        Write-Host "Updating $($corobj.distinguishedName)   Attribute: $attribname" -ForegroundColor Green
                    }
                }
                catch
                {
                    Write-Warning "Error updating $($corobj.distinguishedName)   Attribute: $attribname! Attribute Not Set! Error: $($Error[0])"
                }
            }
        }
    }

    # ---------------------------------------------------------------------------------------------------
    function setLinkedAttributes ($targetOU, $user, $srcMbxAttributes, $srcDomain)
    # ---------------------------------------------------------------------------------------------------
    {
        setLinkedAttribute "altRecipient" "altRecipientBL" $targetOU $user $srcMbxAttributes $srcDomain
        if ($user.properties.contains("altRecipient") -and $srcMbxAttributes.contains("deliverAndRedirect"))
        {
            $user.put("deliverAndRedirect", "$($srcMbxAttributes.item('deliverAndRedirect'))".toupper())
        }
        
        setLinkedAttribute "Manager" "directReports" $targetOU $user $srcMbxAttributes $srcDomain
        
        setLinkedAttribute "publicDelegates" "publicDelegatesBL"  $targetOU $user $srcMbxAttributes $srcDomain
        
        setLinkedAttribute "member" "memberOf"  $targetOU $user $srcMbxAttributes $srcDomain

        setLinkedAttribute "managedBy" "managedObjects"  $targetOU $user $srcMbxAttributes $srcDomain

        setLinkedAttribute "msExchCoManagedByLink" "msExchCoManagedObjectsBL"  $targetOU $user $srcMbxAttributes $srcDomain
    }

    # ---------------------------------------------------------------------------------------------------
    function copyLinkedMailboxTypeAttributes ($user, $srcMbxAttributes)
    # ---------------------------------------------------------------------------------------------------
    {
        $copyAttributes = @()
        $valuedAttributes = @{ }
        
        $accountenable = ($srcMbxAttributes.Item("UserAccountControl").tostring() -band 0x2) -eq 0
        if (-not $accountenable -and (MasterAccountSidIsSelf $srcMbxAttributes))
        {
            $valuedAttributes["msExchRecipientDisplayType"] = $user.properties.Item("msExchRecipientDisplayType").value -bor 2
        }
        else
        {
            $valuedAttributes["msExchRecipientDisplayType"] = 0xC0000006
            if ($srcMbxAttributes.Contains("msExchMasterAccountSid"))
            {
                $copyAttributes += "msExchMasterAccountSid"
            }
            elseif ($srcMbxAttributes.Contains("objectSid"))
            {
                $valuedAttributes["msExchMasterAccountSid"] = $srcMbxAttributes.Item("objectSid").Value
            }
            #this can also be done by carefully arrange "msExchMasterAccountSid" and "objectSid"
            #in the list, avoid the trouble of nested branching. but it's not worth the maintainence effort
        }
        
        [void](copyIfExist $user $copyAttributes $srcMbxAttributes)
    
        foreach($att in $valuedAttributes.getenumerator())
        {
            Write-Verbose "Setting $($att.key) to $($att.value)"
            [void]($user.put($att.key, $att.value))
        }
    }

    # ---------------------------------------------------------------------------------------------------
    function copySpecialMailboxTypeAttributes ($user, $srcMbxAttributes)
    # ---------------------------------------------------------------------------------------------------
    {
        $copyAttributes = "msExchResourceCapacity",
                          "msExchResourceDisplay",
                          "msExchResourceMetaData",
                          "msExchResourceSearchProperties"
                          
        $valuedAttributes = @{ }

        $isResource = $false;

        if ($srcMbxAttributes.Contains("msExchRecipientTypeDetails"))
        {
            $ROOMMAILBOX      = 16
            $EQUIPMENTMAILBOX = 32

            $typedetail = $user.ConvertLargeIntegerToInt64($srcMbxAttributes.Item("msExchRecipientTypeDetails").Value)
            if (($typedetail -band $ROOMMAILBOX) -ne 0)
            {
                $isResource = $true;
                $valuedAttributes["msExchRecipientDisplayType"] = 0x80000706
            }
            elseif (($typedetail -band $EQUIPMENTMAILBOX) -ne 0)
            {
                $isResource = $true;
                $valuedAttributes["msExchRecipientDisplayType"] = 0x80000806
            }
        }

        if ((-not $isResource) -and ($srcMbxAttributes.Contains("msExchRecipientDisplayType")))
        {
            $ROOMMAILBOX      = 0x80000706;
            $EQUIPMENTMAILBOX = 0x80000806;

            $recipientDisplayType = $srcMbxAttributes.Item("msExchRecipientDisplayType").Value;
            if ($recipientDisplayType -eq $ROOMMAILBOX)
            {
                $isResource = $true;
                $valuedAttributes["msExchRecipientDisplayType"] = 0x80000706;
            }
            elseif ($recipientDisplayType -eq $EQUIPMENTMAILBOX)
            {
                $isResource = $true;
                $valuedAttributes["msExchRecipientDisplayType"] = 0x80000806;
            }
        }

        if ($isResource)
        {
            [void](copyIfExist $user $copyAttributes $srcMbxAttributes)

            foreach($att in $valuedAttributes.getenumerator())
            {
                Write-Verbose "Setting $($att.key) to $($att.value)"
                [void]($user.put($att.key, $att.value))
            }
        }
    }

    # ---------------------------------------------------------------------------------------------------
    function copyTeamMailboxAttributes ($targetOU, $user, $srcMbxAttributes, $srcDomain)
    # ---------------------------------------------------------------------------------------------------
    {
        $copyAttributes = "msExchTeamMailboxSharePointUrl",
                          "msExchTeamMailboxExpiration"
        
        $isTeamMailbox = $false;
        
        if ($srcMbxAttributes.Contains("msExchRecipientTypeDetails"))
        {
            $TEAMMAILBOX      = 0x2000000000

            $typedetail = $user.ConvertLargeIntegerToInt64($srcMbxAttributes.Item("msExchRecipientTypeDetails").Value)
            if(($typedetail -band $TEAMMAILBOX) -ne 0)
            {
                $isTeamMailbox = $true;
            }
        }

        if ((-not $isTeamMailbox) -and ($srcMbxAttributes.Contains("msExchRecipientDisplayType")))
        {
            $TEAMMAILBOX      = 0x00000010;

            $recipientDisplayType = $srcMbxAttributes.Item("msExchRecipientDisplayType").Value;
            if ($recipientDisplayType -eq $TEAMMAILBOX)
            {
                $isTeamMailbox = $true;
            }
        }
        
        if($isTeamMailbox)
        {
            Write-Verbose "Setting msExchRecipientDisplayType to 0x80001006."
            [void]($user.put("msExchRecipientDisplayType", 0x80001006))
            [void](copyIfExist $user $copyAttributes $srcMbxAttributes)
            setLinkedAttribute "msExchDelegateListLink" "msExchDelegateListBL"  $targetOU $user $srcMbxAttributes $srcDomain
            setLinkedAttribute "msExchTeamMailboxOwners" ""  $targetOU $user $srcMbxAttributes $srcDomain
        }
    }
    
    # ---------------------------------------------------------------------------------------------------
    function createMEUAndCopyAttrs ($localDC, $localOU, $srcDC, $srcObject)
    # ---------------------------------------------------------------------------------------------------
    {
        $srcAttributes = $srcObject.properties
        $newuser = createMailUserAccount $localDC $localOU $srcAttributes

        #mandatory attributes are all set. go with optional attributes
        copyGalySyncAttributes $newuser $srcAttributes
        copyE2k7OptionalAttributes $newuser $srcAttributes
        setLinkedAttributes $localdc $newuser $srcAttributes $srcdc
        copySpecialMailboxTypeAttributes $newuser $srcAttributes
        copyTeamMailboxAttributes $localdc $newuser $srcAttributes $srcdc

        if ($LinkedMailUser)
        {
            copyLinkedMailboxTypeAttributes $newuser $srcAttributes
        }

        try
        {
            [void]($newuser.SetInfo())
        }
        catch
        {
            Write-Error "Fail to update attributes of local MEU $($newuser.DistinguishedName). $($Error[0])"
            return
        }

        [void](disableEmailAddressPolicy($newuser))
        [void](generateLegacyExchangeDN($newuser))

        #syncback Legacy Exchange DN
        if ($newuser.properties.Contains("LegacyExchangeDN"))
        {
            $X500proxyAddr = "x500:" + $newuser.properties.Item("LegacyExchangeDN")
            mergeLegacyDNToProxyAddress $srcObject ([array]$X500proxyAddr) $true "Source"
        }

        $Global:movecount++
        "Preparation for $Identity done."
    }
    
    # ---------------------------------------------------------------------------------------------------
    function forceMergeObject ($recipienttype, $localOU, $localusr, $localDC, $srcObject, $srcDC)
    # ---------------------------------------------------------------------------------------------------
    {
        $copyAttributes = "msExchMailboxGUID",
                          "msExchArchiveGUID",
                          "msExchArchiveName"

        $X500proxyAddrsToUpdateSrcObj = @()
        $localUserOriginalLegacyDNToX500 = $null
        if ($localusr.Properties.Contains("LegacyExchangeDN"))
        {
            # Original user's legacyDN will be merged to source object.
            $localUserOriginalLegacyDNToX500 = "x500:" + $localusr.Properties.Item("LegacyExchangeDN")
            $X500proxyAddrsToUpdateSrcObj += $localUserOriginalLegacyDNToX500
        }

        $userToUpdate = $null;

        if ($recipienttype -eq 'MailUser')
        {
            [void](disableEmailAddressPolicy($localusr))

            Write-Verbose "Merging Mailbox properties to local MailUser"
            [void](copyIfExist $localusr $copyAttributes $srcObject.properties)
            [void]($localusr.put("msExchVersion", "44220983382016"))
            $logindisabled = ($srcObject.userAccountControl.Value -band 0x2) -ne 0 #AccountDisabled
            if ($LinkedMailUser -and $logindisabled)
            {
                copyLinkedMailboxTypeAttributes $localusr $srcObject.properties
            }

            # Set the proxyAddresses first before merging legacyDN.
            if ($OverwriteLocalObject)
            {
                [void](copyIfExist $localusr "proxyAddresses" $srcObject.properties)
            }

            # Merge source object's legacyDN to local user's proxyAddresses.
            if ($srcObject.properties.Contains("LegacyExchangeDN"))
            {
                $X500proxyAddr = "x500:" + $srcObject.properties.Item("LegacyExchangeDN")
                mergeLegacyDNToProxyAddress $localusr ([array]$X500proxyAddr) $false "Local"
            }

            try
            {
                [void]($localusr.SetInfo())  # Might get Access Denied.
            }
            catch
            {
                Write-Error "Error updating local MEU($($localusr.DistinguishedName)) attributes from source object. Please fix the error and run this script again. Error: $($Error[0])"
            }

            $userToUpdate = $localusr;
        }
        elseif ($recipienttype -eq 'MailContact')
        {
            Write-Verbose "Creating MailUser with same attributes as local MailContact"
            
            $srcMbxAttributes = $srcObject.Properties
            $ContactAttributes = $localusr.Properties

            $newcn = generateNewCN $localDC $srcMbxAttributes.Item("cn").value $localou.distinguishedname
            $newuser = $localOU.create("user", "cn=$newcn")
            
            copyMandatoryAttributes $newuser $ContactAttributes
            copyGalySyncAttributes $newuser $ContactAttributes
            copyE2k7OptionalAttributes $newuser $ContactAttributes

            if ($LinkedMailUser)
            {
                # Copy necessary values (like msExchMasterAccountSid) from contact first, it will be overridden
                # by following one from source mbx.
                copyLinkedMailboxTypeAttributes $newuser $ContactAttributes
            }

            # This must come after copyLinkedMailboxTypeAttributes since it will initialize the value.
            copySpecialMailboxTypeAttributes $newuser $ContactAttributes
            copyTeamMailboxAttributes $localdc $newuser $srcMbxAttributes $srcdc

            [void](copyIfExist $newuser "targetAddress" $ContactAttributes)

            $srcProxyAddresses = $srcMbxAttributes.Item("proxyAddresses")
            $writeErrorIfTargetAddrNotExist = (-not $ContactAttributes.Contains("targetAddress"))
            $srcMbxTargetAddress = getTargetAddress $srcProxyAddresses $writeErrorIfTargetAddrNotExist
            if ($srcMbxTargetAddress -ne $null)
            {
                [void]$newuser.put("targetAddress", $srcMbxTargetAddress)
            }

            # Create the User first, so that we could update the link/backlink of the user before deleting local user.
            try
            {
                [void]($newuser.setinfo())
            }
            catch
            {
                # Terminiate the script if fail to create the new MEU.
                throw "Error creating MailUser cn=$newcn,$($localOU.distinguishedName) with with attributes from source MBX($($srcObject.distinguishedName)). Error: $($Error[0])"
            }
            $newuser.RefreshCache([array]"distinguishedName")
            Write-Host -ForegroundColor green "New MEU $($newuser.distinguishedname) created successfully."

            # backlink needs to be set after user is creaed correctly.
            setLinkedAttributes $localdc $newuser $ContactAttributes $localdc

            Write-Host -ForegroundColor red "Deleteing $($localusr.distinguishedname)"
            try
            {
                deleteLocalUser($localusr)
            }
            catch
            {
                # Terminate the script if fail to delete the object.
                throw "Error deleting local MailContact($($localusr.distinguishedname)). Error: $($Error[0])"
            }

            Write-Verbose "Updating MailUser with with attributes from source MBX($($srcObject.distinguishedName))."

            if ($LinkedMailUser)
            {
                copyLinkedMailboxTypeAttributes $newuser $srcMbxAttributes
            }

            # Set the proxyAddresses first before merging legacyDN.
            if ($OverwriteLocalObject)
            {
                [void](copyIfExist $newuser "proxyAddresses" $srcMbxAttributes)
            }

            $copyAttributes += "sAMAccountName",
                               "userPrincipalName"
            [void](copyIfExist $newuser $copyAttributes $srcMbxAttributes)
            [void]($newuser.put("msExchVersion", "44220983382016"))

            # Merge source object and original local user's LegacyDN to NEW local user's proxyAddresses.
            $X500proxyAddrsToUpdateLocalObj = @()
            if ($localUserOriginalLegacyDNToX500 -ne $null)
            {
                $X500proxyAddrsToUpdateLocalObj += $localUserOriginalLegacyDNToX500
            }
            if ($srcObject.properties.Contains("LegacyExchangeDN"))
            {
                $srcObjectLegacyDNX500 = "x500:" + $srcObject.properties.Item("LegacyExchangeDN")
                if (($localUserOriginalLegacyDNToX500 -ne $null) -and ($localUserOriginalLegacyDNToX500.ToString().ToUpper() -ne $srcObjectLegacyDNX500.ToString().ToUpper()))
                {
                    $X500proxyAddrsToUpdateLocalObj += $srcObjectLegacyDNX500
                }
            }
            mergeLegacyDNToProxyAddress $newuser $X500proxyAddrsToUpdateLocalObj $false "Local"

            [System.Threading.Thread]::Sleep(500)

            try
            {
                [void]($newuser.setinfo())
            }
            catch
            {
                Write-Error "Error updating MailUser cn=$newcn,$($localOU.distinguishedName). Error: $($Error[0])"
            }

            [void](disableEmailAddressPolicy($newuser))
            [void](generateLegacyExchangeDN($newuser))

            # new user's legacyDN will be merged to source object's proxyAddresses
            if ($newuser.properties.Contains("LegacyExchangeDN"))
            {
                $newUserLegacyDNX500 = "x500:" + $newuser.properties.Item("LegacyExchangeDN")
                if (($localUserOriginalLegacyDNToX500 -ne $null) -and ($localUserOriginalLegacyDNToX500.ToString().ToUpper() -ne $newUserLegacyDNX500.ToString().ToUpper()))
                {
                    $X500proxyAddrsToUpdateSrcObj += $newUserLegacyDNX500
                }
            }

            $userToUpdate = $newuser;
        }

        mergeLegacyDNToProxyAddress $srcObject $X500proxyAddrsToUpdateSrcObj $true "Source"

        # Overwrite local object from source object if -OverwriteLocalObject is specified.
        if ($OverwriteLocalObject)
        {
            Write-Verbose "OverwriteLocalObject specified. Updating MailUser with with attributes from source MBX($($srcObject.distinguishedName))."

            $srcAttributes = $srcObject.properties
            copyBasicAttributes $userToUpdate $srcAttributes
            copyGalySyncAttributes $userToUpdate $srcAttributes
            copyE2k7OptionalAttributes $userToUpdate $srcAttributes
            setLinkedAttributes $localdc $userToUpdate $srcAttributes $srcdc
            copySpecialMailboxTypeAttributes $userToUpdate $srcAttributes
            copyTeamMailboxAttributes $localdc $userToUpdate $srcAttributes $srcdc

            try
            {
                [void]($userToUpdate.setinfo())
            }
            catch
            {
                Write-Error "Error updating MailUser($($userToUpdate.DistinguishedName)). Error: $($Error[0])"
            }
        }

        "Preparation for $Identity done. Local recipient info Merged."
        $Global:movecount++
    }
    
    # ---------------------------------------------------------------------------------------------------
    # Delete specified local user through S.DS.P instead of ADSI. This is in order to fix bug #270965
    function deleteLocalUser ($localusr)
    # ---------------------------------------------------------------------------------------------------
    {
        $server = getMailContactOriginatingServer $localusr
        $cnx = createLocalLdapConnection $server
        $deleteRequest = [System.DirectoryServices.Protocols.DeleteRequest]($localusr.distinguishedName.ToString())
        [void]($cnx.SendRequest($deleteRequest))
        $cnx.Dispose()
    }

    # ---------------------------------------------------------------------------------------------------
    function getMailContactOriginatingServer($localusr)
    # ---------------------------------------------------------------------------------------------------
    {
        $mailContact = Get-MailContact $localusr.distinguishedName.Value @DomainControllerParameterSet
        $mailContactServer = $mailContact.OriginatingServer
        $server = $LocalForestDomainController
        if (($mailContactServer -ne $null) -and (-not [string]::IsNullOrEmpty($mailContactServer.ToString())))
        {
            $server = $mailContactServer.ToString()
        }
        Write-Verbose "Local MailContact's OriginatingServer is $server"
        return $server
    }

    # ---------------------------------------------------------------------------------------------------
    # Create a LdapConnection to the local server.
    function createLocalLdapConnection($server)
    # ---------------------------------------------------------------------------------------------------
    {
        $directoryId = [System.DirectoryServices.Protocols.LdapDirectoryIdentifier]($server)
        if ($LocalForestCredential -ne $null)
        {
            $cnx = new-object "System.DirectoryServices.Protocols.LdapConnection"($directoryId, $LocalForestCredential.GetNetworkCredential())
        }
        else
        {
            $cnx = new-object "System.DirectoryServices.Protocols.LdapConnection"($directoryId)
        }
            
        $cnx.SessionOptions.AutoReconnect = $true
        $cnx.SessionOptions.ReferralChasing = [System.DirectoryServices.Protocols.ReferralChasingOptions]::None
        $cnx.SessionOptions.Signing = $true
        $cnx.AuthType = [System.DirectoryServices.Protocols.AuthType]::Kerberos
        $cnx.Bind()
        
        return $cnx
    }
    
#=========================================================================================================
#                                         Initialize code
#=========================================================================================================
    [void]([System.Reflection.Assembly]::LoadWithPartialName("System.DirectoryServices.Protocols"))

    if ($OverwriteLocalObject -and (-not $UseLocalObject))
    {
        throw "Please specify -UseLocalObject when you specify -OverwriteLocalObject parameter."
    }

    $usr = $RemoteForestCredential.UserName
    $pwd = [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($RemoteForestCredential.Password))
    
    if ($LocalForestCredential -ne $null)
    {
        $localusr = $LocalForestCredential.UserName
        $localpwd = [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($LocalForestCredential.Password))
    }
    $Global:movecount = 0

    $srcdc   = New-Object DirectoryServices.DirectoryEntry("LDAP://$RemoteForestDomainController", $usr, $pwd)
    $DomainControllerParameterSet = @{}
    if ($srcdc.guid -eq $null)
    {
        #guid not present, consider src unavailable
        throw "Source Domain controller unavailable or authentication failed."
    }

    try {
        if ($LocalForestCredential -eq $null -and [string]::IsNullorEmpty($LocalForestDomainController))
        {
            $localdc = [ADSI]""
        }
        elseif ($LocalForestCredential -ne $null -and $LocalForestDomainController -ne $null)
        {
            $localdc = New-Object DirectoryServices.DirectoryEntry("LDAP://$LocalForestDomainController", $localusr, $localpwd)
            $DomainControllerParameterSet = @{ DomainController=$LocalForestDomainController; Credential=$LocalForestCredential }
        }
        else
        {
            throw "LocalForestCredential and LocalForestDomainController need to be specified at the same time"
        }

        $escapedtargetou = $null
        $filterObjectClass = $null
        if ([string]::IsNullOrEmpty($TargetMailUserOU))
        {
            # By default, the target OU is in Users container
            $escapedtargetou = "Users"
            $filterObjectClass = "Container"
        }
        else
        {
            $escapedtargetou = getEscapedldapFilterStr $TargetMailUserOU
            $filterObjectClass = "organizationalUnit"
        }
        $OUfilter = "(& (ObjectClass=$filterObjectClass)" +
                    "   (| (name=$escapedtargetou)" +
                    "      (distinguishedname=$escapedtargetou)))"
        $localOU =  findADObject $localdc $OUfilter
        if ($localOU -eq $null)
        {
            throw "Cannot find specified OU or Container: $TargetMailUserOU"
        }
    }
    catch
    {
        throw "Error looking up local OU, Error Msg: $($Error[0])"
    }
}

process
{    
    $escapedIdentity = getEscapedldapFilterStr $Identity
    $filterDN =   "(& (objectClass=user)(!objectClass=computer)" +
                  "   (distinguishedName=$escapedIdentity))"

    $filterParm = "(& (objectClass=user)(!objectClass=computer)" +
                  "   ( (| (mailnickname=$escapedIdentity)" + 
                  "        (cn=$escapedIdentity)" +
                  "        (proxyAddresses=SMTP:$escapedIdentity)" +
                  "        (proxyAddresses=smtp:$escapedIdentity)" +
                  "        (proxyAddresses=X500:$escapedIdentity)" +
                  "        (proxyAddresses=x500:$escapedIdentity)" +
                  "        (objectGUID=$escapedIdentity)" +
                  "        (displayname=$escapedIdentity))))"

    try
    {
        $srcObject = findADObject $srcdc $filterParm

        if ($srcObject -eq $null)
        {
            $srcObject = findADObject $srcdc $filterDN
        
            if ($srcObject -eq $null)
            {
                Write-Error "Error looking up source MBX $identity in source forest."
                return
            }
        }
    }
    catch
    {
        Write-Error "Faile to lookup the source object. Error $($Error[0])"
        return
    }
    
    if (-not $srcObject.properties.contains("mailNickName") -or -not $srcObject.properties.contains("msExchHomeServerName"))
    {
        Write-Error "Source Object $($srcObject.distinguishedName) found, but it is not a Mailbox!."
        return
    }
    
    $accountenable = ($srcObject.properties.Item("UserAccountControl").tostring() -band 0x2) -eq 0
    
    if (-not $accountenable -and -not $srcObject.properties.contains("msExchMasterAccountSid"))
    {
        Write-Error "Source Mailbox is invalid because it is disabled but did not set msExchMasterAccountSid."
        return
    }
    
    try
    {
        $localusr = findLocalObject $localdc $srcObject
    }
    catch
    {
        Write-Error "Error processing $identity, Mailbox not ready to move! Error message: $($error[0])"
        return
    }
    if ($localusr -eq $null)
    {
        try
        {
            #local recipient not exist, source object found, proceed the MEU creation process
            createMEUAndCopyAttrs $localdc $localOU $srcDC $srcObject
        }
        catch
        {
            Write-Error "Error while creating MEU. Error:$($Error[0])";
            return
        }
    }
    else
    {
        Write-Verbose "Local ad account with dupplicate proxy addresses found: $($localusr.distinguishedName)"
        try
        {
            $recipienttype = (get-recipient $localusr.distinguishedname.value @DomainControllerParameterSet).RecipientType
            if ($recipienttype -eq 'MailUniversalDistributionGroup' -or $recipienttype -eq 'UserMailbox' -or $recipienttype -eq 'DynamicDistributionGroup')
            {
                    write-error "Cannot create mail enabled user because an existing mailbox user or mail enabled group already has the same proxy addresses/MasterAccountSid."
            }
            elseif ($recipienttype -eq 'MailUser' -or $recipienttype -eq 'MailContact')
            {
                if ($UseLocalObject)
                {
                    forceMergeObject $recipienttype $localOU $localusr $localDC $srcObject $srcDC
                }
                else
                {
                    write-error ("Cannot create mail enabled user because an existing mail enabled user " +
                                 "or contact already has the same proxy addresses/MasterAccountSid. Please rerun the script with " + 
                                 "'-UseLocalObject' if you want to convert the existing email enabled user or contact to " +
                                 "a mail enabled user that is ready for online mailbox move.")
                }
            }
            else
            {
                write-error "Cannot create mail enabled user because an existing object with type $recipienttype already has the same proxy addresses/MasterAccountSid."
            }
        }
        catch
        {
            Write-Error "Fail to prepare local existed (same Proxyaddress or Masteraccoutsid) user $($localusr.distinguishedName) for move request. Error: $($Error[0])"
            return
        }
    }
}

end
{
    Write-Host -ForegroundColor Black -BackgroundColor Green "$movecount mailbox(s) ready to move."
}


# SIG # Begin signature block
# MIIaxwYJKoZIhvcNAQcCoIIauDCCGrQCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUFsouhYuaefFyBhzoZME3ACTH
# y9CgghWCMIIEwzCCA6ugAwIBAgITMwAAAIliDZ6V02FrqAAAAAAAiTANBgkqhkiG
# 9w0BAQUFADB3MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4G
# A1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSEw
# HwYDVQQDExhNaWNyb3NvZnQgVGltZS1TdGFtcCBQQ0EwHhcNMTUxMDA3MTgxNDAx
# WhcNMTcwMTA3MTgxNDAxWjCBszELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hp
# bmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jw
# b3JhdGlvbjENMAsGA1UECxMETU9QUjEnMCUGA1UECxMebkNpcGhlciBEU0UgRVNO
# Ojk4RkQtQzYxRS1FNjQxMSUwIwYDVQQDExxNaWNyb3NvZnQgVGltZS1TdGFtcCBT
# ZXJ2aWNlMIIBIjANBgkqhkiG9w0BAQEFAAOCAQ8AMIIBCgKCAQEAiqp7T2zEl7js
# zoDvTfJqFbOWhzTKu+q/nUx7DdDt/kVvE8KWox66th8n+O6yqPx8hZ8oGbW8591g
# wDRk1XRsOmPXIm0zQN2qTYXPQJL8uwroK/2Rdj5TtsS1SJOLVmyVhLOVfIMtB99T
# XvIzhHGBeFpk4vZoCjBdW2E8+FsxX8SpeZ0H39yocJv584VTCqTy7wAKZlCRAflX
# qVi9MIq9AF9i2JtPgvSdmr+RjSTjoBi7Zbj825pcGtQeSA4t7akW2ZakGqPHQ7OM
# dKfX3wUgQgBb/UP+582bo7GIPMBXg0eN+GaezzjJeb3c4dGBfGdULyX1IdPbsKzz
# w3VlOODKpwIDAQABo4IBCTCCAQUwHQYDVR0OBBYEFOEoLyVNLefnHPQVNMj6gcN7
# 9bVDMB8GA1UdIwQYMBaAFCM0+NlSRnAK7UD7dvuzK7DDNbMPMFQGA1UdHwRNMEsw
# SaBHoEWGQ2h0dHA6Ly9jcmwubWljcm9zb2Z0LmNvbS9wa2kvY3JsL3Byb2R1Y3Rz
# L01pY3Jvc29mdFRpbWVTdGFtcFBDQS5jcmwwWAYIKwYBBQUHAQEETDBKMEgGCCsG
# AQUFBzAChjxodHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20vcGtpL2NlcnRzL01pY3Jv
# c29mdFRpbWVTdGFtcFBDQS5jcnQwEwYDVR0lBAwwCgYIKwYBBQUHAwgwDQYJKoZI
# hvcNAQEFBQADggEBAEEtDA0RS25aDgdeTf+2Kiy3NxEJyGvynoi0SjwHjTMA0uuy
# 7R5bTBSNJH1WWdPGMkLtlP4UYNcFA3L6uBnbuSufF7g6UdvcShB66IHl1pa0VIUQ
# eithmSVYPzZNbaUmDONRH9AAvmgbMmChVkN+MYx/NITUlzWIrz3LDv/REx5jM47h
# q01rGm7S6iE2nH3nbq6Mp0ItWEOarkK7UAKLmYYCl6Vgz+IY/eF11Kw+BxZiQght
# B0p6IEjcgRZ1ONKyHGBpLVzp/hQUe1hvQFXq20uhe/tTFgra9gYakdypSzr1m7RW
# V7mCFDaVo6Oz+wjwULngfucZC9KdwhBzjEIqxOIwggTsMIID1KADAgECAhMzAAAB
# Cix5rtd5e6asAAEAAAEKMA0GCSqGSIb3DQEBBQUAMHkxCzAJBgNVBAYTAlVTMRMw
# EQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVN
# aWNyb3NvZnQgQ29ycG9yYXRpb24xIzAhBgNVBAMTGk1pY3Jvc29mdCBDb2RlIFNp
# Z25pbmcgUENBMB4XDTE1MDYwNDE3NDI0NVoXDTE2MDkwNDE3NDI0NVowgYMxCzAJ
# BgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25k
# MR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xDTALBgNVBAsTBE1PUFIx
# HjAcBgNVBAMTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjCCASIwDQYJKoZIhvcNAQEB
# BQADggEPADCCAQoCggEBAJL8bza74QO5KNZG0aJhuqVG+2MWPi75R9LH7O3HmbEm
# UXW92swPBhQRpGwZnsBfTVSJ5E1Q2I3NoWGldxOaHKftDXT3p1Z56Cj3U9KxemPg
# 9ZSXt+zZR/hsPfMliLO8CsUEp458hUh2HGFGqhnEemKLwcI1qvtYb8VjC5NJMIEb
# e99/fE+0R21feByvtveWE1LvudFNOeVz3khOPBSqlw05zItR4VzRO/COZ+owYKlN
# Wp1DvdsjusAP10sQnZxN8FGihKrknKc91qPvChhIqPqxTqWYDku/8BTzAMiwSNZb
# /jjXiREtBbpDAk8iAJYlrX01boRoqyAYOCj+HKIQsaUCAwEAAaOCAWAwggFcMBMG
# A1UdJQQMMAoGCCsGAQUFBwMDMB0GA1UdDgQWBBSJ/gox6ibN5m3HkZG5lIyiGGE3
# NDBRBgNVHREESjBIpEYwRDENMAsGA1UECxMETU9QUjEzMDEGA1UEBRMqMzE1OTUr
# MDQwNzkzNTAtMTZmYS00YzYwLWI2YmYtOWQyYjFjZDA1OTg0MB8GA1UdIwQYMBaA
# FMsR6MrStBZYAck3LjMWFrlMmgofMFYGA1UdHwRPME0wS6BJoEeGRWh0dHA6Ly9j
# cmwubWljcm9zb2Z0LmNvbS9wa2kvY3JsL3Byb2R1Y3RzL01pY0NvZFNpZ1BDQV8w
# OC0zMS0yMDEwLmNybDBaBggrBgEFBQcBAQROMEwwSgYIKwYBBQUHMAKGPmh0dHA6
# Ly93d3cubWljcm9zb2Z0LmNvbS9wa2kvY2VydHMvTWljQ29kU2lnUENBXzA4LTMx
# LTIwMTAuY3J0MA0GCSqGSIb3DQEBBQUAA4IBAQCmqFOR3zsB/mFdBlrrZvAM2PfZ
# hNMAUQ4Q0aTRFyjnjDM4K9hDxgOLdeszkvSp4mf9AtulHU5DRV0bSePgTxbwfo/w
# iBHKgq2k+6apX/WXYMh7xL98m2ntH4LB8c2OeEti9dcNHNdTEtaWUu81vRmOoECT
# oQqlLRacwkZ0COvb9NilSTZUEhFVA7N7FvtH/vto/MBFXOI/Enkzou+Cxd5AGQfu
# FcUKm1kFQanQl56BngNb/ErjGi4FrFBHL4z6edgeIPgF+ylrGBT6cgS3C6eaZOwR
# XU9FSY0pGi370LYJU180lOAWxLnqczXoV+/h6xbDGMcGszvPYYTitkSJlKOGMIIF
# vDCCA6SgAwIBAgIKYTMmGgAAAAAAMTANBgkqhkiG9w0BAQUFADBfMRMwEQYKCZIm
# iZPyLGQBGRYDY29tMRkwFwYKCZImiZPyLGQBGRYJbWljcm9zb2Z0MS0wKwYDVQQD
# EyRNaWNyb3NvZnQgUm9vdCBDZXJ0aWZpY2F0ZSBBdXRob3JpdHkwHhcNMTAwODMx
# MjIxOTMyWhcNMjAwODMxMjIyOTMyWjB5MQswCQYDVQQGEwJVUzETMBEGA1UECBMK
# V2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0
# IENvcnBvcmF0aW9uMSMwIQYDVQQDExpNaWNyb3NvZnQgQ29kZSBTaWduaW5nIFBD
# QTCCASIwDQYJKoZIhvcNAQEBBQADggEPADCCAQoCggEBALJyWVwZMGS/HZpgICBC
# mXZTbD4b1m/My/Hqa/6XFhDg3zp0gxq3L6Ay7P/ewkJOI9VyANs1VwqJyq4gSfTw
# aKxNS42lvXlLcZtHB9r9Jd+ddYjPqnNEf9eB2/O98jakyVxF3K+tPeAoaJcap6Vy
# c1bxF5Tk/TWUcqDWdl8ed0WDhTgW0HNbBbpnUo2lsmkv2hkL/pJ0KeJ2L1TdFDBZ
# +NKNYv3LyV9GMVC5JxPkQDDPcikQKCLHN049oDI9kM2hOAaFXE5WgigqBTK3S9dP
# Y+fSLWLxRT3nrAgA9kahntFbjCZT6HqqSvJGzzc8OJ60d1ylF56NyxGPVjzBrAlf
# A9MCAwEAAaOCAV4wggFaMA8GA1UdEwEB/wQFMAMBAf8wHQYDVR0OBBYEFMsR6MrS
# tBZYAck3LjMWFrlMmgofMAsGA1UdDwQEAwIBhjASBgkrBgEEAYI3FQEEBQIDAQAB
# MCMGCSsGAQQBgjcVAgQWBBT90TFO0yaKleGYYDuoMW+mPLzYLTAZBgkrBgEEAYI3
# FAIEDB4KAFMAdQBiAEMAQTAfBgNVHSMEGDAWgBQOrIJgQFYnl+UlE/wq4QpTlVnk
# pDBQBgNVHR8ESTBHMEWgQ6BBhj9odHRwOi8vY3JsLm1pY3Jvc29mdC5jb20vcGtp
# L2NybC9wcm9kdWN0cy9taWNyb3NvZnRyb290Y2VydC5jcmwwVAYIKwYBBQUHAQEE
# SDBGMEQGCCsGAQUFBzAChjhodHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20vcGtpL2Nl
# cnRzL01pY3Jvc29mdFJvb3RDZXJ0LmNydDANBgkqhkiG9w0BAQUFAAOCAgEAWTk+
# fyZGr+tvQLEytWrrDi9uqEn361917Uw7LddDrQv+y+ktMaMjzHxQmIAhXaw9L0y6
# oqhWnONwu7i0+Hm1SXL3PupBf8rhDBdpy6WcIC36C1DEVs0t40rSvHDnqA2iA6VW
# 4LiKS1fylUKc8fPv7uOGHzQ8uFaa8FMjhSqkghyT4pQHHfLiTviMocroE6WRTsgb
# 0o9ylSpxbZsa+BzwU9ZnzCL/XB3Nooy9J7J5Y1ZEolHN+emjWFbdmwJFRC9f9Nqu
# 1IIybvyklRPk62nnqaIsvsgrEA5ljpnb9aL6EiYJZTiU8XofSrvR4Vbo0HiWGFzJ
# NRZf3ZMdSY4tvq00RBzuEBUaAF3dNVshzpjHCe6FDoxPbQ4TTj18KUicctHzbMrB
# 7HCjV5JXfZSNoBtIA1r3z6NnCnSlNu0tLxfI5nI3EvRvsTxngvlSso0zFmUeDord
# EN5k9G/ORtTTF+l5xAS00/ss3x+KnqwK+xMnQK3k+eGpf0a7B2BHZWBATrBC7E7t
# s3Z52Ao0CW0cgDEf4g5U3eWh++VHEK1kmP9QFi58vwUheuKVQSdpw5OPlcmN2Jsh
# rg1cnPCiroZogwxqLbt2awAdlq3yFnv2FoMkuYjPaqhHMS+a3ONxPdcAfmJH0c6I
# ybgY+g5yjcGjPa8CQGr/aZuW4hCoELQ3UAjWwz0wggYHMIID76ADAgECAgphFmg0
# AAAAAAAcMA0GCSqGSIb3DQEBBQUAMF8xEzARBgoJkiaJk/IsZAEZFgNjb20xGTAX
# BgoJkiaJk/IsZAEZFgltaWNyb3NvZnQxLTArBgNVBAMTJE1pY3Jvc29mdCBSb290
# IENlcnRpZmljYXRlIEF1dGhvcml0eTAeFw0wNzA0MDMxMjUzMDlaFw0yMTA0MDMx
# MzAzMDlaMHcxCzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYD
# VQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xITAf
# BgNVBAMTGE1pY3Jvc29mdCBUaW1lLVN0YW1wIFBDQTCCASIwDQYJKoZIhvcNAQEB
# BQADggEPADCCAQoCggEBAJ+hbLHf20iSKnxrLhnhveLjxZlRI1Ctzt0YTiQP7tGn
# 0UytdDAgEesH1VSVFUmUG0KSrphcMCbaAGvoe73siQcP9w4EmPCJzB/LMySHnfL0
# Zxws/HvniB3q506jocEjU8qN+kXPCdBer9CwQgSi+aZsk2fXKNxGU7CG0OUoRi4n
# rIZPVVIM5AMs+2qQkDBuh/NZMJ36ftaXs+ghl3740hPzCLdTbVK0RZCfSABKR2YR
# JylmqJfk0waBSqL5hKcRRxQJgp+E7VV4/gGaHVAIhQAQMEbtt94jRrvELVSfrx54
# QTF3zJvfO4OToWECtR0Nsfz3m7IBziJLVP/5BcPCIAsCAwEAAaOCAaswggGnMA8G
# A1UdEwEB/wQFMAMBAf8wHQYDVR0OBBYEFCM0+NlSRnAK7UD7dvuzK7DDNbMPMAsG
# A1UdDwQEAwIBhjAQBgkrBgEEAYI3FQEEAwIBADCBmAYDVR0jBIGQMIGNgBQOrIJg
# QFYnl+UlE/wq4QpTlVnkpKFjpGEwXzETMBEGCgmSJomT8ixkARkWA2NvbTEZMBcG
# CgmSJomT8ixkARkWCW1pY3Jvc29mdDEtMCsGA1UEAxMkTWljcm9zb2Z0IFJvb3Qg
# Q2VydGlmaWNhdGUgQXV0aG9yaXR5ghB5rRahSqClrUxzWPQHEy5lMFAGA1UdHwRJ
# MEcwRaBDoEGGP2h0dHA6Ly9jcmwubWljcm9zb2Z0LmNvbS9wa2kvY3JsL3Byb2R1
# Y3RzL21pY3Jvc29mdHJvb3RjZXJ0LmNybDBUBggrBgEFBQcBAQRIMEYwRAYIKwYB
# BQUHMAKGOGh0dHA6Ly93d3cubWljcm9zb2Z0LmNvbS9wa2kvY2VydHMvTWljcm9z
# b2Z0Um9vdENlcnQuY3J0MBMGA1UdJQQMMAoGCCsGAQUFBwMIMA0GCSqGSIb3DQEB
# BQUAA4ICAQAQl4rDXANENt3ptK132855UU0BsS50cVttDBOrzr57j7gu1BKijG1i
# uFcCy04gE1CZ3XpA4le7r1iaHOEdAYasu3jyi9DsOwHu4r6PCgXIjUji8FMV3U+r
# kuTnjWrVgMHmlPIGL4UD6ZEqJCJw+/b85HiZLg33B+JwvBhOnY5rCnKVuKE5nGct
# xVEO6mJcPxaYiyA/4gcaMvnMMUp2MT0rcgvI6nA9/4UKE9/CCmGO8Ne4F+tOi3/F
# NSteo7/rvH0LQnvUU3Ih7jDKu3hlXFsBFwoUDtLaFJj1PLlmWLMtL+f5hYbMUVbo
# nXCUbKw5TNT2eb+qGHpiKe+imyk0BncaYsk9Hm0fgvALxyy7z0Oz5fnsfbXjpKh0
# NbhOxXEjEiZ2CzxSjHFaRkMUvLOzsE1nyJ9C/4B5IYCeFTBm6EISXhrIniIh0EPp
# K+m79EjMLNTYMoBMJipIJF9a6lbvpt6Znco6b72BJ3QGEe52Ib+bgsEnVLaxaj2J
# oXZhtG6hE6a/qkfwEm/9ijJssv7fUciMI8lmvZ0dhxJkAj0tr1mPuOQh5bWwymO0
# eFQF1EEuUKyUsKV4q7OglnUa2ZKHE3UiLzKoCG6gW4wlv6DvhMoh1useT8ma7kng
# 9wFlb4kLfchpyOZu6qeXzjEp/w7FW1zYTRuh2Povnj8uVRZryROj/TGCBK8wggSr
# AgEBMIGQMHkxCzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYD
# VQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xIzAh
# BgNVBAMTGk1pY3Jvc29mdCBDb2RlIFNpZ25pbmcgUENBAhMzAAABCix5rtd5e6as
# AAEAAAEKMAkGBSsOAwIaBQCggcgwGQYJKoZIhvcNAQkDMQwGCisGAQQBgjcCAQQw
# HAYKKwYBBAGCNwIBCzEOMAwGCisGAQQBgjcCARUwIwYJKoZIhvcNAQkEMRYEFATH
# XIgDWwXZH7tAcSXEdqvWi9cnMGgGCisGAQQBgjcCAQwxWjBYoDCALgBQAHIAZQBw
# AGEAcgBlAC0ATQBvAHYAZQBSAGUAcQB1AGUAcwB0AC4AcABzADGhJIAiaHR0cDov
# L3d3dy5taWNyb3NvZnQuY29tL2V4Y2hhbmdlIDANBgkqhkiG9w0BAQEFAASCAQAt
# tIrXFwJZAjUruRCTkETnWNdqYZj/twFgJ2vR/rztpGKL9Kg2EeOJtuhQmCApWcPt
# 4G4S1V/LWJWLNpPXsVZWaKQ+KvyvYqknRbZoFmyNlYEGFhhqIWLJ+7l1JrA034zE
# Vsd+XuAE//ewranpayEnG1qp/sp9rknbk/dJAgruxoRQIZ3/4xgDUODxtv5oMvUa
# f/B28FLvKFmXa6DhboB4JuC4aOEAGwi/f4GQ8VZOwESbAOM3WnxJ/mn59N5h0eDO
# V71bmMFT75jbloBfVTtB7AXwFVhZTxxpElSsrRjlz+zA/LbdvmxLGp63dg6xdWoh
# 2eR5CkWBqwpj8LRAHz33oYICKDCCAiQGCSqGSIb3DQEJBjGCAhUwggIRAgEBMIGO
# MHcxCzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdS
# ZWRtb25kMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xITAfBgNVBAMT
# GE1pY3Jvc29mdCBUaW1lLVN0YW1wIFBDQQITMwAAAIliDZ6V02FrqAAAAAAAiTAJ
# BgUrDgMCGgUAoF0wGAYJKoZIhvcNAQkDMQsGCSqGSIb3DQEHATAcBgkqhkiG9w0B
# CQUxDxcNMTYwMTA3MDA0ODM3WjAjBgkqhkiG9w0BCQQxFgQUdhkhBsOAztN/9/X8
# 1xOAJgZJC94wDQYJKoZIhvcNAQEFBQAEggEAOaAYRlz/ZPjduOtfKPwzzTpWK+eM
# YYOsPksQcBJiV6xWIOhfFETL5XwTzY4GAwel5r4/svAxicc0kZwyf3PXWYazq7EV
# jQQ61ZdR1dbABEKfHIrONBQXNLQOYiVnk5IdqKWIUoK8TciGgYMM82T+5ScIPd7C
# 4XJQOMYUDxzrDeJgMesnjBgFoxNfAvFB6zv56Zzgj3rQl63hCGX/8EWDXy+wR3+u
# 1xB83LJuKO0oZXmxRxYmXJEb+Ct3l3856D+XlN4m7/CU8weOT3XoEFXXw/TG1Gql
# NLclWbLIT2LKDonGS48I71Yg6CgikMKGIlZVAoC+XE9KOSNVRe/fLcyb2w==
# SIG # End signature block
