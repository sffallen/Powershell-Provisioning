#region globals
$hrpath = "\\racqgroup\data\TEC-Business Collaboration\Field Services\Desktop & Field Services\User Provisioning\HR\Current\HRUP02.CSV"
$hrdata = get-content -Path  $hrpath | Select-Object -skip 1 | ConvertFrom-Csv
$datadate = Get-date -Format "dd/MM/yy"

$leftTab = "" #holds the current tabpage on the LHS
$rightTab = "" #holds the current tabpage on the RHS

$memberToggle = @{}
$objLoaded = @{}
#$objLoadedClass = @{}

$userProps = "*"
$groupProps = @("mail","description","info","memberof","members","objectClass")
$PcProps = @("description", "operatingsystem", "operatingsystemversion", "ipv4address", "canonicalname", "enabled", "LastLogonDate", "memberof","objectClass")
$AdObjProps = @{ "User" = $userProps; "Group" = $groupProps; "Computer" = $PcProps }

Function dn2sam([string]$dn) { return $dn.Substring(3,($dn.indexof(","))-3) } 
$AllGroups = (Get-ADGroup -filter *) | %  { dn2sam $_ } 

$searchHint = "Enter the name of an AD object or search for a user's name or an AD group name.`n`n"
$searchHint += "    AD Group searches will find any AD group with the entered text somewhere in the name. `n"
$searchHint += "    &&  can be used as a wildcard to search AD groups  e.g   'sl-file&&PMO' "
$searchHint += "`n`n    User name searches will use   |   to separate the first and last name`n e.g 'Bee Bee|'   or   'Jan|Van der B' "

$textview = [System.Drawing.Image]::FromFile("\\racqgroup\Data\TEC-Business Collaboration\Field Services\Desktop & Field Services\Icons\bluelist32.png")
$listview = [System.Drawing.Image]::FromFile("\\racqgroup\Data\TEC-Business Collaboration\Field Services\Desktop & Field Services\Icons\blueplay32.png")
$memberOfPic = [System.Drawing.Image]::FromFile("\\racqgroup\data\TEC-Business Collaboration\Field Services\Desktop & Field Services\Icons\one-to-many2.png")
$memberPic = [System.Drawing.Image]::FromFile("\\racqgroup\data\TEC-Business Collaboration\Field Services\Desktop & Field Services\Icons\many-to-one.png")

$sky = [System.Drawing.Color]::FromArgb(255,225,255,255) #colour for read only fields
$lemon = [System.Drawing.Color]::FromArgb(255,255,249,235) #colour for listbox cells
$grey = [System.Drawing.Color]::FromArgb(255,220,220,220) 
$console = [System.Drawing.Color]::FromArgb(255,240,240,240) 
$softPink = [System.Drawing.Color]::FromArgb(255,255,225,225) 
$warnPink = [System.Drawing.Color]::FromArgb(255,255,200,200) 
$warnRed = [System.Drawing.Color]::FromArgb(255,255,100,100)
$warnOrange = [System.Drawing.Color]::FromArgb(255,255,150,100)

$listboxDrawMode = {

        param([object]$s, [System.Windows.Forms.DrawItemEventArgs]$e)

        if ($e.Index -gt -1)
            {
                <# If the item is selected set the background color to SystemColors.Highlight 
                 or else set the color to either WhiteSmoke or White depending if the item index is even or odd #>
                $color = if(($e.State -band [System.Windows.Forms.DrawItemState]::Selected) -eq [System.Windows.Forms.DrawItemState]::Selected){ 
                    [System.Drawing.SystemColors]::Highlight
                }else{
                    if($e.Index % 2 -eq 0){
                        [System.Drawing.Color]::WhiteSmoke
                    }else{
                        [System.Drawing.Color]::White
                    }
                }

                # Background item brush
                $backgroundBrush = New-Object System.Drawing.SolidBrush $color
                # Text color brush
                $textBrush = New-Object System.Drawing.SolidBrush $e.ForeColor

                # Draw the background
                $e.Graphics.FillRectangle($backgroundBrush, $e.Bounds)
                # Draw the text
                $e.Graphics.DrawString($s.Items[$e.Index], $e.Font, $textBrush, $e.Bounds.Left, $e.Bounds.Top, [System.Drawing.StringFormat]::GenericDefault)

                # Clean up
                $backgroundBrush.Dispose()
                $textBrush.Dispose()
            }
            $e.DrawFocusRectangle()
}

#endregion globals

#region Called Functions

#region App-Specific

#region Logical

Function ParseInput($str, $tbp){
    if ($str.length -lt 3) { $lblReady[$tbp].Text = "Search strings must contain at least 3 characters." }
    else { 
        exitReady -tbp $tbp       
        
        $infoText = ""; $objArray = @(); $user = ""; $pc = ""; $grp = ""; $obj=$null
                        
        if($str.contains("|")) {$objArray = SearchUsers $str; enterMatches -tbp $tbp -objArray $objArray}
        elseif($str.contains("&")) {$objArray = searchGroups $str; enterMatches -tbp $tbp -objArray $objArray}
        else {        
         
            $obj = getObject -objID $str -props $AdObjProps
            if ($obj) {             
                $objLoaded[$tbp] = $obj
                $infoText = getObjInfo $obj    
                enterInfo -tbp $tbp -infoText $infoText -obj $obj            
            }
            else
            {     
                if(! $str.contains(" ")) {$objArray = (SearchUsers ($str + "|")) + (SearchUsers ("|" + $str)) }
                else {$objArray = SearchUsers $str}
                $objArray += searchGroups $str             
                enterMatches -tbp $tbp -objArray $objArray 
            } 
        }       
    }  
}

Function SearchUsers($names) {
            if ($names -like '*|*') { $split = $names.split("{'|'}") }
            else { $split = $names.split("{' '}") }
        	$name1 = "*" + $split[0] + "*"
	        $name2 = "*" + $split[$split.Length-1] + "*" 
            if ($name1 -ne "**") {	            
                if ($name2 -ne "**") { $list1 = Get-ADUser -Filter {(givenname -Like $name1) -and (surname -like $name2) } }
                else { $list1 = Get-ADUser -Filter {(givenname -Like $name1)} }
            }
            else { if ($name2 -ne "**") {$list1 = Get-ADUser -Filter {(surname -like $name2)} } }          
	        if ($name1 -or $name2) { 
                $list2 = @()               
                $list1 | % { $list2 += $_.samAccountname + ":     "+$_.givenname + " " + $_.surname+"`r" }
                return $list2
            }
}

Function SearchGroups($names) {
 	$names = $names.Split("{&}")
	$Matches = $global:AllGroups
	foreach ($name in $names) { $Matches = $Matches|where-object {$_ -match $name} }
	$Matches = $matches| Sort 
    return $matches
 }

 Function GetObjInfo( $obj ) {
    if ($obj.ObjectClass -eq "User") { return getUserInfo $obj } 
    if ($obj.ObjectClass -eq "Group") { return getGroupInfo $obj } 
    if ($obj.ObjectClass -eq "Computer") { return getPCInfo $obj } 
 }

 Function GetUserInfo($user) { #uses getsecondaryaccounts
 
    $contactarray = @( $user.telephoneNumber,$user.otherTelephone, $user.OfficePhone, $user.MobilePhone, $user.HomePhone, $user.ipphone )
    $contactarray = $contactarray | unique
    $contacts = ""
    $contactarray | % { if ($_) { $contacts +=  [string]$_.trim() + "     " } }
    
    $displayText = $user.samaccountname + "   "+ $user.givenName + " " + $user.surname + "   "+((getsecondaryaccounts $user.samaccountname) -join(" ")) + "`n`n"       
    $displayText += $user.office +"   "+$user.mail+"    "+$contacts+ "`n"
    if ($user.lockedout) { $displaytext += "* ACCOUNT LOCKED *   " }
    if ($user.passwordexpired) { $displaytext += "* PASSWORD EXPIRED *   " }
    if ($user.enabled) { $displaytext += "Enabled   " }
    else { $displaytext += "* DISABLED *   " }
    $displaytext += $user.canonicalname + "`n"
    $displaytext += $user.division + "  -  " + $user.department + "  -  " + $user.title + "`n"
    return $displayText
} 

 Function GetPcInfo($pc) { 
    
    $displayText = $pc.name + "`n"    
    $displayText += $pc.description + "`n"
    $displayText += "Last logon:     " + [string]$pc.lastlogondate + "`n"
    $displayText += "OS:               "+$pc.operatingsystem + "`n"
    $displayText += "Version:         "+$pc.operatingsystemversion + "`n" 
    $displayText += $pc.ipv4address + "`n" 
    $displayText += $pc.canonicalname + "`n" 
    if ($pc.enabled) { $displayText += "Enabled `n" } else  { $displayText += "Disabled `n" }
    return $displaytext
 } 

 Function GetGroupInfo($group) { 

    $displayText = $group.samaccountname + "`n`n"
    $displayText += $group.mail + "`n"
    $displayText += $group.description + "`n`n"
    $displayText += $group.info + "`n"
    return $DisplayText
 } 

 Function initTabPages($tbp) {
    
    $memberToggle[$tbp] = "Groups"
 }

 Function DisplayMembership($tbp) {

    Cleartext $rtbMembership[$tbp] 
    $objLoaded[$tbp] = getObject -objID $objLoaded[$tbp].name -props $AdObjProps
    if ($memberToggle[$tbp] -eq "Groups") { $objLoaded[$tbp].Memberof | % { (dn2sam $_) + "`n" } | sort | % { $rtbMembership[$tbp].text += $_ } }
    else { $objLoaded[$tbp].Members | % { (dn2sam $_) + "`n" } | sort | % { $rtbMembership[$tbp].text += $_ } }

    $lblMembership[$tbp].text = $memberToggle[$tbp] + " for " + $objLoaded[$tbp].name
 }

 Function DisplayComparison {    
    $left = $rtbMembership[$global:leftTab].text.split("`n")    
    $right = $rtbMembership[$global:rightTab].text.split("`n")

    foreach ($item in $left) {
        $rtbEitherLeftOrRight[$global:rightTab].Text += $item + "`n"
        if($item -in $right) { $rtbBothLeftandRight[$global:leftTab].Text += $item + "`n" }
        else { $rtbLeftOnly[$global:leftTab].Text += $item + "`n" }
    }
    foreach ($item in $right) { if($item -notin $left) { $rtbRightOnly[$global:rightTab].Text += $item + "`n" } }

    $lblLeftOnly[$global:leftTab].text = $objLoaded[$global:leftTab].name + " only" 
    $lblBothLeftandRight[$global:leftTab].text = "Both " + $objLoaded[$global:leftTab].name + " and " + $objLoaded[$global:rightTab].name
    $lblEitherLeftorRight[$global:rightTab].text = "Either " + $objLoaded[$global:leftTab].name + " or " + $objLoaded[$global:rightTab].name
    $lblRightOnly[$global:rightTab].text = $objLoaded[$global:rightTab].name + " only" 
 }

 Function enact ($tbp) {
     
    if ( $global:memberToggle[$tbp] -eq "Members" ) { write-host " Changing members not yet supported "; exitEnact -tbp $tbp }
    else {        

        if ($btnEnact[$tbp].Text -eq "Add") { 
            $rtbAddText = $rtbRightHalf[$tbp].Text            
            addGroups -obj $objLoaded[$tbp] -groups $rtbAddText.split("`n") -outTxt $rtbGroupResults[$tbp] -ref $lblRefSet[$tbp].Text
        }
        if ($btnEnact[$tbp].Text -eq "Add + Remove") {
            $rtbAddText = $rtbRightThird[$tbp].Text 
            $rtbRemText = $rtbMidThird[$tbp].Text 
            addGroups -obj $objLoaded[$tbp] -groups $rtbAddText.split("`n") -outTxt $rtbGroupResults[$tbp] -ref $lblRefSet[$tbp].Text
            removeGroups -obj $objLoaded[$tbp] -groups $rtbRemText.split("`n") -outTxt $rtbGroupResults[$tbp] -ref $lblRefSet[$tbp].Text -append $True
        }
            if ($btnEnact[$tbp].Text -eq "Remove") { 
            $rtbRemText = $rtbRightHalf[$tbp].Text 
            removeGroups -obj $objLoaded[$tbp] -groups $rtbRemText.split("`n") -outTxt $rtbGroupResults[$tbp] -ref $lblRefSet[$tbp].Text
        }
        enterMembership -tbp $tbp -obj $objLoaded[$tbp]
        enterGroupResults -tbp $tbp
        exitEnact -tbp $tbp
        
    }

 }

 Function MovePdUsers ($correctPD,$users,$copyGroups,$removeGroups,$rtbOut) {  # uses getPD, isinAD

    $createmsg = ""
    $usersEmpty = $False
    $inPD = $True
    $copyinremove = $True
    $addedperms = @()

    if ($users) { #checks all users match the inputted correct PD
        $users | % { $inPD = $inPD -and ( (getPD $_) -eq $correctPD) }
    }
    else {$usersEmpty = $True}
    if ($inPD) { #creates the correct PD if it does not exist 
        $created = isInAD -id $correctPD -objtype "group"
        if (-not $created) { $createmsg = createPD ([array]$users)[0] }
         
        #add all users to the correct PD
        try { Add-ADGroupMember $correctPD -Members $users } catch {}
                
        #checks copy groups are also in the removegroups 
        $copyInRemove = $True
        if ($copygroups) {            
            $copygroups | % { $copyInRemove = $copyInRemove -and ( $_ -in $removeGroups ) }

            if ($copyInRemove) {              

                #add access in copyGroups to correctPD
                foreach($cpygrp in $copyGroups) {
                    $perms = getPermissions $cpygrp
                    foreach ($prm in $perms) {
                         try {Add-ADGroupMember $prm -Members $correctPD; $addedPerms += $prm} catch {}                         
                    }
                }
            }
        }
        #remove users from removeGroups
        if ($copyinremove) {
            foreach($user in $users) {
                    foreach($rmvgrp in $removeGroups){
                        try {Remove-ADGroupMember $rmvgrp –Member $user -Confirm:$false} catch{}
                    }
                }
        } # end if $copygroups
    } # end if $inPD
    
    #output

    $out = ""
    $addedperms = $addedperms | sort | unique    
    if ($usersEmpty) { $out = "No users detected: No action taken." }
    elseif (-not $inpd) { 
        $out = "Not all the given users belong to the given correct PD: No action taken`n" 
        $out += "Given PD: " + $correctPD + "`n"
        $users | % { $out += $_ + " :  " + (getPD $_) + "`n" }
    }
    elseif ($copygroups -and (-not $copyInRemove)) { 
        $out = "Not all given copy groups were in the remove groups. No action taken `n"
    }
    else {
        $out = "Processed as per KB0012961.`nCorrected for:`n`n Correct PD: "+$correctPD
        if ($createmsg) {$out += "`n`n" + $createmsg}
        $out += "`n`n Users who have been added to the correct PD (if needed) and removed from any other PD Groups:`n"
        $users | % { $out += $_ + "`n" }
        if ($copygroups) {
            $out += "`n Incorrect PDs that have had access copied to "+$correctPD+":`n"
            $copygroups | % { $out += $_ + "`n" }            
        }
        if ($removeGroups) {
           $out += "`n Incorrect PDs that the above users were removed from (where needed): `n" 
           $removegroups | % { $out += $_ + "`n" }
        }
        if ($addedperms) {
                $out += "`n Access added (if not already present) to "+$correctPD+":`n"
                $addedperms | % { $out += $_ + "`n" }
        } 
    } # end else
    $rtbout.text = $out
}

 Function ListsMatch ($listoflists) {
        $first = $listoflists[0]
        $success = $true
        foreach ($listn in $listoflists[1..$listoflists.count]) {
            if ($listn.count -ne $first.count) { $success = $false }
            $listn | % { if (!($first -contains $_)) { $success = $false } }              
        }
        return $success
    }   

 Function RenamePD ($newPDtbx,$oldPDtbx,$userbox,$outbox) { # uses listsmatch renamePDTo RTB2List dn2sam
        $newPD = $newPDtbx.text
        $oldPD = $oldPDtbx.text
        $newusernames = RTB2List $userbox
        $oldmembers = $(get-adgroup -Properties members $oldPD).members | % { dn2sam $_ }
        if (ListsMatch @($newusernames,$oldmembers) ) {
            $result = RenamePDTo -new $newPD -old $oldPD
            if($result -eq "Done") {
                $out = "Processed as per KB0012961.`nCorrected for:`n`n Correct PD: "
                $out += $newPD + "`n`n "+$newpd+" has been renamed from "+$oldpd
                $out += "`n`n Renaming will have conserved all access previously assigned to "+$oldPD
                $out += "`n`nUser who are members of this PD are: `n"
                $newusernames | % { $out += "`n" + $_ }
                $outbox.text = $out
            } else {$outbox.Text = $result}
        } else {
            $outbox.text = $newpd + "`n members and `n" + $oldPD + "`nmembers do not match. `n"
            $count = 1
            $newusernames | % { $outbox.text += ($count.ToString() + ": " + $_ + "`n");$count+=1 } 
        }
 } 

 Function RenamePDto ($new, $old) {    #returns outcome string
        $email = (parseEmail $new) + "@racq.com.au"
        $group = get-adgroup $old -Properties mail,cn,samaccountname
        $dupe =  get-adobject -filter {cn -eq $new}
        $obj = get-adobject -filter {cn -eq $old}
        if ($dupe) {
            return "Object with name: "+ $new + " already exists!"
        } else {
            if (!$obj.count) {
                rename-adobject $obj -newname $new   
                set-adgroup $group.samaccountname -samaccountname $new -displayname $new -Replace @{mail = $email}
                for  ($count=1; $count -le 100; $count++) 
                { 
                    $newgroup = get-adgroup -filter {cn -eq $new} -Properties mail,cn,samaccountname;
                    if ($newgroup) {break}
                    write-host $count
                    start-sleep -milliseconds 50
                }
                return "Done"
            } else {
                $objstring = ""
                $obj | % {$objstring += $_.distinguishedname} 
                return "Multiple groups with the CN: " + $old + "`n" + $objsting
            }
        }
                          
    }


 #endregion Logical

 #region State

 Function exitReady ($tbp) {

    hide $pnlReady[$tbp]  
    $lblReady[$tbp].text = $searchHint 
 
 }

 Function enterReady ($tbp) { 
     
    show $pnlReady[$tbp]
    readWrite $tbxReady[$tbp]
    clearText $tbxReady[$tbp]
    $lblReady[$tbp].text = $searchHint  
    $tbxReady[$tbp].focus()
    $form.AcceptButton = $btnReady[$tbp]
    
    exitSetActions $tbp
    exitInfo $tbp  
    exitMatches $tbp 
    exitCompare 
	exitGroupresults $tbp    
 }

 Function enterInfo ($tbp, $infoText,$obj) {

    show @( $pnlInfo[$tbp], $pnlRef[$tbp], $pnlMembership[$tbp], $pnlEnact[$tbp] )
    enterMembership -tbp $tbp -obj $obj

    $lbxMatches[$tbp].Items.Clear()
    $lbxMembership[$tbp].Items.Clear()
    $lblMembership[$tbp].text = "Groups for " + $objLoaded[$tbp].name
    $rtbInfo[$tbp].Text = $infoText 
    
    hide $pnlMatches[$tbp]
    $tbxRef[$tbp].focus()
    $form.AcceptButton = $btnResetInfo[$tbp]
 }

 Function exitInfo ($tbp) {
    ClearText $rtbInfo[$tbp]
    Hide @( $pnlInfo[$tbp], $pnlRef[$tbp], $pnlEnact[$tbp], $btnCancelEnact[$tbp], $btnEnact[$tbp] )
    
    exitMembership ($tbp)
 }

 Function enterMatches ($tbp, $objArray) {
    $objArray | sort | % { $lbxMatches[$tbp].Items.Add($_) }
    $pnlMatches[$tbp].Visible = $True 
    $form.AcceptButton = $btnResetMatches[$tbp]
 }

 Function exitMatches ($tbp) {
    $lbxMatches[$tbp].Items.Clear()
    ClearText $rtbMatches[$tbp]
    Hide $pnlMatches[$tbp]
 }

 Function enterSetActions ($tbp) {
    
    if ($tbxRef[$tbp].Text.length -gt 2) {
        $lblRefSet[$tbp].Text = $tbxRef[$tbp].Text
        clearText $tbxRef[$tbp]    
        hide @( $pnlRef[$tbp] )
        show @( $pnlSetActions[$tbp], $btnCancelEnact[$tbp] )
    }
 }

 Function exitSetActions ($tbp) {
    show @($pnlRef[$tbp])
    hide @($pnlSetActions[$tbp], $btnCloneGroups[$tbp], $btnCancelEnact[$tbp], $btnEnact[$tbp], $lblAddGroups[$tbp], $lblRemGroups[$tbp] )
    clearText $lblRefSet[$tbp]
 }

 Function enterMembership ($tbp, $obj) {

    $objLoaded[$tbp] = $obj

    if ($obj.objectClass -eq "Group") { show $btnToggleMembershipDirection[$tbp] }
    if ($obj.objectClass -eq "User") { show $btnELV[$tbp] }
    if ($lbxMembership[$tbp].Visible) {
        toggleView -lbx $lbxMembership[$tbp] -rtb $rtbMembership[$tbp] -btn $btnToggleMembershipView[$tbp] -lbxPic $listview -rtbPic $textview 
    }

    displayMembership -tbp $tbp   

    if ($tbp -like "Main*") {$global:leftTab =  $tbp}
    if ($tbp -like "Aux*") {$global:rightTab = $tbp}
    if ($global:leftTab -and $global:rightTab) { enterCompare }
    
 }

Function exitMembership ($tbp) {

    clearText $rtbMembership[$tbp]
    Hide @($pnlMembership[$tbp], $btnToggleMembershipDirection[$tbp],$btnELV[$tbp])

    if ($tbp -like "Main*") {$global:leftTab =  $null}
    if ($tbp -like "Aux*") {$global:rightTab = $null}

    $objLoaded[$tbp] = $null
    if ($memberToggle[$tbp] -eq "Members" ) { toggleMembershipDirection -tbp $tbp -update $false }
}

Function enterEnact ($tbp, $action, $addgroups = "", $remgroups = "") {
    
    
    if ($action -eq "Add") {
        $rtbRightHalf[$tbp].text = $addgroups
        show @( $rtbRightHalf[$tbp], $lblAddGroups[$tbp] )
        $grpSize = 240
        $color = $warnOrange

        
    }
    if ($action -eq "Remove") {
        $rtbRightHalf[$tbp].text = $remgroups
        show @( $rtbRightHalf[$tbp], $lblRemGroups[$tbp] )
        $drawpoint.x = 320
        $drawpoint.y = 0
        $lblRemGroups[$tbp].Location=$drawpoint 
        $grpSize = 240
        $color = $warnPink
    }
    if ($action -eq "Add + Remove") {
        $rtbMidThird[$tbp].text = $remgroups
        $rtbRightThird[$tbp].text = $addgroups
        show @( $rtbMidThird[$tbp], $rtbRightThird[$tbp], $lblAddGroups[$tbp], $lblRemGroups[$tbp] )
        $grpSize = 160
        $color = $warnRed
    }

    $rtbMembership[$tbp].Width = $grpSize
    show $btnEnact[$tbp]
    hide @($lblMembership[$tbp], $pnlSetActions[$tbp] )
    $btnEnact[$tbp].Text = $action
    $btnEnact[$tbp].BackColor = $color
    $form.AcceptButton = $null
}

Function exitEnact ($tbp) {
    exitSetActions -tbp $tbp 
    $rtbMembership[$tbp].Width = 480
    hide @( $rtbRightHalf[$tbp], $rtbMidThird[$tbp], $rtbRightThird[$tbp] )
    clearText @( $rtbRightHalf[$tbp], $rtbMidThird[$tbp], $rtbRightThird[$tbp] )
    show @( $lblMembership[$tbp] )
    $drawpoint.x = 160
    $drawpoint.y = 0
    $lblRemGroups[$tbp].Location=$drawpoint 
}

Function enterCompare {
    displayComparison
    $pnlCompare[$global:leftTab].Visible = $true
    $pnlCompare[$global:rightTab].Visible = $true    
} 

Function exitCompare {

    foreach( $tab in $tabs.keys) {
    foreach( $tbp in $tbpList[$tab]) {
        $pnlCompare[$tbp].Visible = $false

        if($tab -eq "Main") { clearText @( $rtbLeftOnly[$tbp], $rtbBothLeftAndRight[$tbp] ) }
        if($tab -eq "Aux") { clearText @( $rtbRightOnly[$tbp], $rtbEitherLeftOrRight[$tbp] ) }
    }
    }
}

Function enterGroupResults($tbp) {
    show $pnlGroupResults[$tbp]
}

Function exitGroupResults($tbp) {
    hide $pnlGroupResults[$tbp]
    clearText $pnlGroupResults[$tbp]
}

Function enterELV($tbp) {
    write-host "ELV"
}

Function toggleMembershipDirection ($tbp, $update = $true) {
       if ($memberToggle[$tbp] -eq "Groups") {
        $memberToggle[$tbp] = "Members"
        $btnToggleMembershipDirection[$tbp].image = $memberPic
    }
    else {
        $memberToggle[$tbp] = "Groups"
        $btnToggleMembershipDirection[$tbp].image = $memberOfPic        
    }
    $btnAddGroups[$tbp].Text = "Add " + $memberToggle[$tbp]
    $btnRemgroups[$tbp].Text = "Remove " + $memberToggle[$tbp]
    $btnAddRemGroups[$tbp].Text = "Add + Remove " + $memberToggle[$tbp]
    $btnClearGroups[$tbp].Text = "Clear " + $memberToggle[$tbp]
    $btnCloneGroups[$tbp].Text = "Clone " + $memberToggle[$tbp]

    $lblRemGroups[$tbp].Text = "Remove " + $memberToggle[$tbp]
    $lblAddGroups[$tbp].Text = "Add " + $memberToggle[$tbp]
    
    if($update) { DisplayMembership -tbp $tbp }  
}



#endregion State

#endregion App-Specific

#region Component Setting
    
Function readOnly ($objs) { $objs | % { $_.readOnly = $True; $_.backColor = $sky } }
Function readWrite ($objs) { $objs | % { $_.readOnly = $False; $_.backColor = "white" } }
Function show ($objs) { $objs | % { $_.Visible = $True } }
Function hide ($objs) { $objs | % { $_.Visible = $False } }
Function clearText ($objs) { $objs | % { $_.Text = "" } }
Function drawAsListbox ($objs) { $objs | % { $_.DrawMode = [System.Windows.Forms.DrawMode]::OwnerDrawFixed; $_.add_DrawItem($listboxDrawMode) } }
Function toggleView ($lbx, $rtb, $btn, $lbxPic, $rtbPic) {
    if($lbx.Visible) {

        show $rtb
        hide $lbx
        $btn.Image = $lbxPic
        $lbx.Items | % { $rtb.text += $_ + "`n" }
        $lbx.Items.Clear()
       
    } else {
        show $lbx
        hide $rtb
        $btn.Image = $rtbPic
        $rtb.text.Split("`n") | % { $lbx.Items.Add($_) }
        clearText $rtb
    }
}


#endregion Component Setting

#region utility

Function getCommon ($ListofLists) {
    if ($ListofLists) {
        $common = $ListofLists[0]
        $ListofLists = $LIstofLists[1..$ListofLists.count]
                 
        while ($ListofLists) {
            $compare = $ListofLists[0]
            $ListofLists = $LIstofLists[1..$ListofLists.count]
            $common = $common | where { $_ -in $compare }
        } 
        return $common
    }
}

Function getNestedPermissions ($ADObjectName,$objtype=$null,$dictusers=$null,$dictgroups=$null,$dictcomputers=$null) {

    if ( (!$objtype) -or (!$allgroups) -or (($objtype -eq "user") -and (!$allusers)) -or (($objtype -eq "user") -and (!$allusers)) ) { 
	    $obj = Get-ADObject -Filter {SamAccountName -eq $ADObjectName}
	    if ($obj.ObjectClass -eq "user") { $raw = $(Get-ADUser $ADObjectName -Properties MemberOf).MemberOf }
	    elseif ($obj.ObjectClass -eq "group") { $raw = $(Get-ADGroup $ADObjectName -Properties MemberOf).MemberOf }
	    else { $raw = $(Get-ADComputer $ADObjectName -Properties MemberOf).MemberOf }
	    
        $groups = @()
	    while ($raw) {
		    $objname = $raw[0]
		    if (!($groups -contains $objname)) {$raw += $(Get-ADGroup $objname -Properties Memberof).MemberOf} 
		    $groups += $objname
		    $raw = $raw[1..$raw.count]
	    }
	    $sams = $groups | % { $(get-adgroup $_ -prop samaccountname).samaccountname }	
    }
    else {
        if ($objtype -eq "user") { $raw = $dictusers[$ADObjectName].memberof }
        if ($objtype -eq "computer") { $raw = $dictcomputers[$ADObjectName].memberof }
        if ($objtype -eq "groups") { $raw = $dictgroups[$ADObjectName].memberof }
               
        $groups = @{}
	    while ($raw) {
            $grpname = $raw[0]
		    if (!($groups[$grpname])) {
                $raw += $dictgroups[(dn2sam $grpname)].memberof
                $groups[$grpname] = $True
            }
            $raw = $raw[1..$raw.count]
        }
        $sams = @()
        $groups.keys | % { $sams += dn2sam $_ }
    }
	return $sams
}

Function getPermissions ($ADObjectName) {
    $obj = Get-ADObject -Filter {SamAccountName -eq $ADObjectName}
    if ($obj.ObjectClass -eq "user") { $dns = $(Get-ADUser $ADObjectName -Properties MemberOf).MemberOf }
	elseif ($obj.ObjectClass -eq "group") { $dns = $(Get-ADGroup $ADObjectName -Properties MemberOf).MemberOf }
	else { $dns = $(Get-ADComputer $ADObjectName -Properties MemberOf).MemberOf }  
    return $dns | % { $(get-adgroup $_ -properties samaccountname).samaccountname }  
}

Function GetSecondaryAccounts ($user) { # uses isInAD
    
    $user = $user.trim() 
    if ($user.length -lt 6) { return @() }
    else {

    $accounts = @()
    
    $eid = $user.substring(1)
    $fifthLast = $user.substring(($user.length -5),1)
    $last4 = $user.substring($user.length -4)
    
    $accounts += "A" + $eid
    $accounts += "S" + $eid
    $accounts += "D" + $eid
    $accounts += "T" + $eid
    $accounts += "TA" + $eid
    if ($fifthLast -gt 0) { foreach ( $n in 1..8) { $accounts += "TN" + $n + $eid } }
    else { foreach ( $n in 1..8) { $accounts += "T0" + $n + $last4 } }
    return $accounts | where { isInAD $_ -objType "user" }  
    }
}

Function isInAD ($id, $objtype) {
    $currentPreference = $ErrorActionPreference
    $ErrorActionPreference = 'SilentlyContinue'
    $success = $true
    
    if ($objtype -eq "group") { try { $obj = get-adgroup $id } catch {$success = $false; write-host "Error in group AD detection" $_.Exception.message} } 
    elseif ($objtype -eq "user") { try { $obj = get-aduser $id } catch {$success = $false; write-host "Error in user AD detection"} } 
    elseif ($objtype -eq "computer") { try { $obj = get-adcomputer $id } catch {$success = $false; write-host "Error in pc AD detection"} } 
    else { $success = $false; write-host "Error in object type detection"}

    $ErrorActionPreference = $currentPreference

    return ($success) 
}

Function GetCurrentUser{ return [Security.Principal.WindowsIdentity]::GetCurrent().Name.ToUpper().Replace("RACQGROUP\","") }

Function LogChanges($ref,$scriptuser,$object,$change,$details) {
	$filepath = "\\racqgroup\data\TEC-Business Collaboration\Field Services\Desktop & Field Services\User Provisioning\Logs\allusers.log"	
    $details = $details.replace("`n","|")	
	$logstring = (get-date -format "yyyy-MM-dd HH:mm:ss") + " " + $scriptuser + " " + $object + " " + $ref + " " + $change + " { " + $details + " }`n" 
    $mutex = new-object System.Threading.Mutex $false,'AllUsersLog'
    $mutex.WaitOne() > $null
	Add-Content  -Path $filepath -Value $logstring -Force
    $mutex.ReleaseMutex()    
}

Function getObject( $objID, $props = @{ "User" = "*"; "Group" = "*"; "Computer"= "*" } ) {
    $obj = get-adobject -filter { (samaccountname -eq $objId) -or (name -eq $objId) -or (displayname -eq $objId) }
    if ($obj) {
        $obj = $obj | ? { $_.objectClass -in @("User","Group","Computer") }     
        if ($obj.objectClass -eq "User") { $retObj = get-adUser $obj.distinguishedname -Properties $props[$obj.objectClass] }
        if ($obj.objectClass -eq "Group") { $retObj = get-adGroup $obj.distinguishedname -Properties $props[$obj.objectClass] }
        if ($obj.objectClass -eq "Computer") { $retObj = get-adComputer $obj.distinguishedname -Properties $props[$obj.objectClass] } 
        return $retobj    
    }
}

Function AddGroups($obj, $groups, $outTxt, $ref) { # uses getNestedPermissions, getCommon, isInAD, ApplyRestrictions, GetCurrentUser, LogChanges

#Todo? check if user group being added to a computer or vice versa

    $unparsed = $groups -split '\r?\n' # convert from text to an array
    $groups0 = @()
    $unparsed | % { $groups0 += $_.TrimEnd() }    

    $groups0 = $groups0 | ? { $_ }
    $groups = $groups0 | ? { (isInAD $_ "group") }
          
    $notinAD = $groups0 | ? { $_ -notin [array]$groups }       
    
    #compare groups-to-add to existing groups

    $objGroups = getNestedPermissions $obj.samaccountname | sort | unique
    $preExisting = @()
	if ($objgroups) { $preExisting = getCommon @($groups, $objgroups) } 
      
    #read restricted database and exclude restricted groups

    $elvdata = get-content -Path "\\racqgroup\Data\TEC-Business Collaboration\Field Services\Desktop & Field Services\User Provisioning\Checks\Elevated Security\Elevated Security Group Listing.csv" | convertfrom-csv
    $unparsed = $elvdata.groupname 
    $allrestricted = @()

    $unparsed | % { $allrestricted += $_.TrimEnd() }

    $allrestricted = $allrestricted | ? { $_ }
    $allrestricted = $allrestricted | ? { get-adgroup $_ }  
    
    $restricted = getCommon @($groups, $allrestricted)

    $denied = ApplyRestrictions $restricted $obj $elvdata
    
    $restricted = $restricted | ? { $_ }
    $restricted = $restricted | ? { get-adgroup $_ }  
    
    $toAdd =  $groups | ? { $_ -notin [array]($preExisting) }
    $toAdd = $toAdd | ? { $_ -notin [array]($denied) }           

    #attempt to add groups - make list of successful/ failed

    $failed = @()
    $added = @()
    
    $toAdd | % {
        $group = $_
        try { Add-ADGroupMember $_ -Members $obj; $added += $_ }
	    catch [system.exception] { $failed += $group; write-host "`r`n" EXCEPTION: $_.Exception.message "`r`n Failed to add "$group "to" $obj "`r`n" }
    }

    #update output

    $details = ""    
    if ($obj.ObjectClass -eq "user") { 
        $user = Get-ADUser $obj.name
        $username = $user.givenName + " " + $user.surname
    } 
    else { $username = "" }
    if ($added) { $details += "Groups Added to "+$obj.name+"  "+$username+": `n"+ ($added -join "`n") + "`n`n" }
    if ($preExisting) { $details += "Groups Not Added Due To Access Already Applied (Possibly Nested): `n"+ ($preExisting -join "`n") + "`n`n" }
    if ($failed) { $details += "Groups Not Added Due To Command Failure: `n"+ ($failed -join "`n")  + "`n`n"  }
    if ($denied) { $details += "Groups Not Added Due To KB0012680 Restriction: `n"+ ($denied -join "`n") + "`n`n"  }
    if ($notInAD) { $details += "Groupnames Not Recognised: `n"+ ($notInAD -join "`n") + "`n`n"  }

    $outTxt.text = $details

    #logchanges

    $ScriptUser = GetCurrentUser
    if ($details) { LogChanges -ref $ref -scriptuser $ScriptUser -object $obj.name -change "Groups Added" -details $details }
       
}

Function RemoveGroups($obj, $groups, $outTxt, $ref, $append=$false) {
    
    $unparsed = $groups -split '\r?\n' # convert from text to an array
    $groups0 = @()
    $unparsed | % { $groups0 += $_.TrimEnd() }    

    $groups0 = $groups0 | ? { $_ }
    $groups = $groups0 | ? { (isInAD $_ "group") }
          
    $notinAD = $groups0 | ? { $_ -notin [array]$groups }       
    
    #compare groups-to-remove to existing groups

    $objGroups = getPermissions $obj.samaccountname | sort | unique
    $toRemove = getCommon @($groups, $objgroups) 

    $notPresent = $groups | ? { $_ -notin [array]($toRemove) }          

    #attempt to remove groups - make list of successful/ failed

    $failed = @()
    $removed = @()
    
    $toRemove | % {
        $group = $_
        try { Remove-ADGroupMember $_ -Members $obj -Confirm:$false; $removed += $_ }
	    catch [system.exception] { $failed += $group; write-host "`r`n" EXCEPTION: $_.Exception.message "`r`n Failed to remove "$group "from" $obj "`r`n" }
    }

    #update output

    $details = ""    
    if ($removed) { $details += "Groups Removed from "+$obj.name+": `n"+ ($removed -join "`n") + "`n`n" }
    if ($notpresent) { $details += "Groups Not Removed Due To Not Being Present: `n"+ ($notPresent -join "`n") + "`n`n" }
    if ($failed) { $details += "Groups Not Removed Due To Command Failure: `n"+ ($failed -join "`n")  + "`n`n"  }    
    if ($notInAD) { $details += "Groupnames Not Recognised: `n"+ ($notInAD -join "`n") + "`n`n"  }

    if( !$append ) {$outTxt.Text = ""}
    $outTxt.text += $details

    #logchanges

    $ScriptUser = GetCurrentUser
    if ($details) { LogChanges -ref $ref -scriptuser $ScriptUser -object $obj.name -change "Groups Removed" -details $details }
}

Function ApplyRestrictions($groups, $obj, $restrictions) {
    $objtype = (get-adobject -filter { name -eq $obj.name}).objectClass
    $denied = @()
    foreach($group in $groups) { 
        $restrictiontype = ( $restrictions | ? { $_.groupname -eq $group } )
        if ($restrictiontype) {
            if ($restrictiontype.type -eq "ELV") { $denied += $group }
            if ($restrictiontype.type -eq "ALRT") { $denied += $group }
            if ( ($restrictiontype.type -eq "ROLE") -and ($objtype -in @("user","computer")) ) { $denied += $group }
            if ( ($restrictiontype.type -eq "USER") -and ($objtype -in @("group","computer")) ) { $denied += $group }
        }
    }
    return $denied
} 

Function RTB2List($rtb,$filter = @('"')) {
    $list1 = $rtb.Text.split("`n")
    $list1 = $list1 | % {$_.TrimEnd()}
    $list1 = $list1 | % {foreach ($removeChar in $filter) { $_.replace($removeChar,'') } }
    $list1 = $list1 | ? {$_} 
    return [array]$list1
} 

Function List2RTB($list1,$rtb,$append = $false) {
    if (-not $append) { $rtb.Text = "" }
    $list1 | % {$rtb.Text += $_ + "`n" }
}

function getPD($eid,$future=$false) { # requires hrdata uses PDFilter, singlespace
    if ($eid[0] -in @('u','U')) { $eid = $eid.substring(1) }
	$hr = $hrdata | where-object{$eid -eq $_.EMP}
    if($future) { $hr = $hrfdata | where-object{$eid -eq $_.EMP} }
	$pd = "SG-"+$hr.division+"-"+$hr.department+"-"+$hr.unit+"-"+$hr."RACQ position title"
    $pd = PDFilter $pd
	if ($pd.Length -gt 64) { $pd = $pd.Substring(0,64) }
    $pd = legalEndChar $pd
	Return $pd.TrimEnd()
}

function getUnit($eid,$future=$false) {
    if ($eid[0] -in @('u','U')) { $eid = $eid.substring(1) }
	$hr = $hrdata | where-object{$eid -eq $_.EMP}
    if($future) { $hr = $hrfdata | where-object{$eid -eq $_.EMP} }
	$pd = "SG-"+$hr.division+"-"+$hr.department+"-"+$hr."unit long"
    $pd = PDFilter $pd
	if ($pd.Length -gt 64) { $pd = $pd.Substring(0,64) }
    $pd = legalEndChar $pd
	Return $pd.TrimEnd()
}

function getDepartment($eid,$future=$false) {
    if ($eid[0] -in @('u','U')) { $eid = $eid.substring(1) }
	$hr = $hrdata | where-object{$eid -eq $_.EMP}
    if($future) { $hr = $hrfdata | where-object{$eid -eq $_.EMP} }
	$pd = "SG-"+$hr.division+"-"+$hr."department long"
    $pd = PDFilter $pd
	if ($pd.Length -gt 64) { $pd = $pd.Substring(0,64) }
    $pd = legalEndChar $pd
	Return $pd.TrimEnd()
}

function getDivision($eid,$future=$false) {
    if ($eid[0] -in @('u','U')) { $eid = $eid.substring(1) }
	$hr = $hrdata | where-object{$eid -eq $_.EMP}
    if($future) { $hr = $hrfdata | where-object{$eid -eq $_.EMP} }
	$pd = "SG-"+$hr."division long"
    $pd = PDFilter $pd
	if ($pd.Length -gt 64) { $pd = $pd.Substring(0,64) }
    $pd = legalEndChar $pd
	Return $pd.TrimEnd()
}

function PDFilter($pd) { # uses singleSpace
    $illegal = @("!","#","$","%","'","*","+","/","=","?","^","_","|",",","~","{","}",'`','"','(',')',',',':',';','<','>','@','[','\',']')
    $illegal | % { $pd = $pd.replace($_," ") }
	Return singleSpace $pd    
}

function singleSpace($str, $endSpace = $False, $startSpace = $False) {
    $str = $str -split '\s+' -join " "
    if (!$endSpace) {$str = $str.trimEnd()}
    if (!$startSpace) {$str = $str.trimStart()}
    return $str
}

Function legalEndChar($pd) { 
    $illegalEndChar = @(".","-","&")
    while ( $illegalEndChar -contains $pd[-1] ) { $pd=$pd.Substring(0,($pd.length-1)) } 
    return $pd
}

function createPD ($user,$future=$false) { # uses isInAD, parseEmail, getPD, getUnit, getDepartment, getDivision
    $msg = ""
    $pd = getpd $user -future $future
    if ($pd -eq "SG----") { 
            $msg += "User " + $user + " not recognised." + $pd + "`n"
            return $null
    } if ( isInAD -objType "group" -id $pd ) {
         $msg += $pd +" is already in AD.`n"
         return $null
    }

    $path = "OU=Role PD Groups,OU=Groups,DC=racqgroup,DC=local"
    $email = (parseEmail $pd) + "@racq.com.au"
    try {New-AdGroup -Name $pd -path $path -Groupscope Global -GroupCategory Security -Description "PD Group" -OtherAttributes @{mail = $email}; $msg+= $pd + " created.`n" }   
    catch {$msg+="Error creating pd group "+$pd+" on path "+$path+" with email "+$email+"`n"+$_.Exception.message+"`n"}
    $unit = getunit $user -future $future
    if ( !(isInAD -id $unit -objType "group") ) {
        $email = (parseEmail $unit) + "@racq.com.au"
        try {New-AdGroup -Name $unit -path $path -Groupscope Global -GroupCategory Security -Description "Unit Group" -OtherAttributes @{mail = $email}; $msg+= $unit+ " created.`n" }   
        catch {$msg+="Error creating unit group "+$unit+" on path "+$path+" with email "+$email+"`n"+$_.Exception.message+"`n"}
        $department = getDepartment $user -future $future
        if ( !(isInAD -id $department -objType "group") ) {
            $email = (parseEmail $department) + "@racq.com.au"
            try {New-AdGroup -Name $department -path $path -Groupscope Global -GroupCategory Security -Description "Department Group" -OtherAttributes @{mail = $email}; $msg+= $department+ " created.`n" }   
            catch {$msg+="Error creating department group "+$department+" on path "+$path+" with email "+$email+"`n"+$_.Exception.message+"`n"}
            $div = getDivision $user -future $future
            if ( !(isInAD -id $div -objectType "group") ) {
                $email = (parseEmail $div) + "@racq.com.au"
                try {New-AdGroup -Name $div -path $path -Groupscope Global -GroupCategory Security -Description "Division Group" -OtherAttributes @{mail = $email}; $msg+= $div+ " created.`n" }   
                catch {$msg+="Error creating division group "+$division+" on path "+$path+" with email "+$email+"`n"+$_.Exception.message+"`n"}
            }
        }
    }
    try { Add-ADGroupMember $unit -Members $pd }
	catch {$msg+="Could not add "+$pd+" to "+$unit+".`n"+$_.Exception.message+"`n"}
    if ($department) {
        try { Add-ADGroupMember $department -Members $unit }
	    catch {$msg+="Could not add "+$unit+" to "+$department+".`n"+$_.Exception.message+"`n"}
    }
    if ($div) {
        try { Add-ADGroupMember $div -Members $department }
	    catch {$msg+="Could not add "+$department+" to "+$div+".`n"+$_.Exception.message+"`n"}
        try { Add-ADGroupMember "SG-All Staff" -Members $div }
	    catch {$msg+="Could not add "+$div+" to SG-All Staff.`n"+$_.Exception.message+"`n"}
    }
    return $msg
}

function parseEmail ($str) { 
    $illegal = @("!","#","$","%","&","'","*","+","/","=","?","^","_","|",",","~","{","}",'`','"','(',')',',',':',';','<','>','@','[','\',']','-','.'," ")
    $illegal | % { $str = $str.replace($_,"") }
    return $str 
}



#endregion utility

#endregion Called Functions

#region Form and Tabs

#region Form

$form = New-Object System.Windows.Forms.Form
$DrawSize = New-Object System.Drawing.Size
$DrawPoint = New-Object System.Drawing.Point	
$form.Name = "form"
$form.Text = "Provisioning Tool"
$DrawSize.Width = 1400
$DrawSize.Height = 800
$form.ClientSize = $DrawSize

#endregion Form

#region Tabs

$tabs = @{}
$tabsAux = New-Object System.Windows.Forms.TabControl
$tabsMain = New-Object System.Windows.Forms.TabControl

$tabs["Main"] = $tabsMain
$tabs["Main"].Name = "tabsMain"
$DrawPoint.X = 0
$DrawPoint.Y = 0
$tabs["Main"].Location = $DrawPoint
$DrawSize.Width = 700
$DrawSize.Height = 800
$tabs["Main"].Size = $Drawsize 
$form.Controls.Add($tabsMain)


$tabs["Aux"] = $tabsAux
$tabs["Aux"].Name = "tabsAux"
$DrawPoint.X = 700
$DrawPoint.Y = 0
$tabs["Aux"].Location = $DrawPoint
$DrawSize.Width = 700
$DrawSize.Height = 800
$tabs["Aux"].Size = $Drawsize 
$form.Controls.Add($tabsAux)

#endregion tabs

#region Tabpages

###### Logic assumes there will be no tabpages on the Main tab with the same name as a tabpage on the Aux tab

$tbps = @{}
$tbpList = @{}
$tbpList["Main"] = [array]@("Main1","Main2")
$tbpList["Aux"] = @("Aux1","Aux2","Aux3")


foreach( $tab in $tabs.keys) {
    foreach( $tbp in $tbpList[$tab]) {
        
        $tbps[$tbp] = New-Object System.Windows.Forms.TabPage
        $DrawPoint.X = 3
        $DrawPoint.Y = 3
        $tbps[$tbp].Name = "tbps[" + $tbp + "]"
        $tbps[$tbp].Location = $Drawpoint
        $tbps[$tbp].Text = $tbp
        $DrawSize.Width = 691
        $DrawSize.Height = 749
        $tbps[$tbp].Size = $DrawSize          
        $tabs[$tab].Controls.Add($tbps[$tbp]) 
        
        initTabPages -tbp $tbp             
    }        
}

#region tbpPdTool
    
    $tbpPdTool = New-Object System.Windows.Forms.TabPage
    $DrawPoint.X = 3
    $DrawPoint.Y = 3
    $tbpPdTool.Name = "tbpPdTool"
    $tbpPdTool.Location = $Drawpoint
    $tbpPdTool.Text = "PD Tool"
    $DrawSize.Width = 691
    $DrawSize.Height = 749
    $tbpPdTool.Size = $DrawSize          
    $tabs["Main"].Controls.Add($tbpPdTool) 

    #region TabsPdTool

    $tabsPdTool = New-Object System.Windows.Forms.TabControl
    $tabsPdTool.Name = "tabsPDTool"
    $DrawPoint.X = 0
    $DrawPoint.Y = 0
    $tabsPdTool.Location = $DrawPoint
    $DrawSize.Width = 680
    $DrawSize.Height = 730
    $tabsPdTool.Size = $Drawsize 
    $tbpPdTool.Controls.Add($tabsPdTool)

    #endregion TabsPdTool

    #region tbpMoveUsers
    
    $tbpMoveUsers = New-Object System.Windows.Forms.TabPage
    $DrawPoint.X = 3
    $DrawPoint.Y = 3
    $tbpMoveUsers.Name = "tbpMoveUsers"
    $tbpMoveUsers.Location = $Drawpoint
    $tbpMoveUsers.Text = "Move Users"
    $DrawSize.Width = 670
    $DrawSize.Height = 420
    $tbpMoveUsers.Size = $DrawSize          
    $tabsPdTool.Controls.Add($tbpMoveUsers) 

    #endregion tbpMoveUsers

    #region tbpRenamePD
    
    $tbpRenamePD = New-Object System.Windows.Forms.TabPage
    $DrawPoint.X = 3
    $DrawPoint.Y = 3
    $tbpRenamePD.Name = "tbpRenamePD"
    $tbpRenamePD.Location = $Drawpoint
    $tbpRenamePD.Text = "Rename PD"
    $DrawSize.Width = 670
    $DrawSize.Height = 420
    $tbpRenamePD.Size = $DrawSize          
    $tabsPdTool.Controls.Add($tbpRenamePD) 

    #endregion tbpRenamePD

#endregion tbpPdTool


#endregion Tabpages

#endregion Form and Tabs

#region Ready

$pnlReady = @{}
$tbxReady = @{}
$btnReady = @{}
$lblReady = @{}

foreach( $tab in $tabs.keys) {
    foreach( $tbp in $tbpList[$tab]) {
        
        #region pnlReady
                
        $pnlReady[$tbp] = New-Object System.Windows.Forms.Panel
        $pnlReady[$tbp].Name = "pnlReady["+$tbp+"]"
        #$pnlReady[$tbp].Visible = $False
        $DrawPoint.X = 0
        $DrawPoint.Y = 0
        $pnlReady[$tbp].Location = $Drawpoint
        $DrawSize.Width = 690
        $DrawSize.Height = 190
        $pnlReady[$tbp].Size = $DrawSize
        #$pnlReady[$tbp].BackColor = $lemon
        $tbps[$tbp].Controls.Add($pnlReady[$tbp])             

        #endregion pnlReady

        #region tbxReady

        $tbxReady[$tbp] = New-Object System.Windows.Forms.TextBox
        $tbxReady[$tbp].Name = "tbxReady["+$tbp+"]"
        $tbxReady[$tbp].Text = ""
        $DrawPoint.X = 20
        $DrawPoint.Y = 10
        $tbxReady[$tbp].Location = $DrawPoint
        $DrawSize.Width = 560
        $DrawSize.Height = 20
        $tbxReady[$tbp].Size = $DrawSize
        $tbxReady[$tbp].add_GotFocus({param($sender,$e)
            tbxReadyOnClick -tbp ($sender.name.split('[')[1].split(']')[0])
        })
        $pnlReady[$tbp].Controls.Add($tbxReady[$tbp])            

        #endregion tbxReady

        #region btnReady

        $btnReady[$tbp] = New-Object System.Windows.Forms.Button
        $btnReady[$tbp].Name = "btnReady["+$tbp+"]"
        $btnReady[$tbp].Text = "Search"
        $DrawPoint.X = 590
        $DrawPoint.Y = 10
        $btnReady[$tbp].Location = $DrawPoint
        $DrawSize.Width = 80
        $DrawSize.Height = 40
        $btnReady[$tbp].Size = $DrawSize        
        $btnReady[$tbp].add_Click({param($sender,$e)
            btnReadyOnClick -tbp ($sender.name.split('[')[1].split(']')[0])
        })
        $pnlReady[$tbp].Controls.Add($btnReady[$tbp])

        #endregion btnReady       

        #region lblReady

        $lblReady[$tbp] = New-Object System.Windows.Forms.Label
        $lblReady[$tbp].Name = "lblReady[$tbp]"
        $lblReady[$tbp].Text = $searchHint
        $lblReady[$tbp].TextAlign = "MiddleCenter" 
        $DrawPoint.X = 20
        $DrawPoint.Y = 30
        $lblReady[$tbp].Location = $DrawPoint
        $DrawSize.Width = 500
        $DrawSize.Height = 120
        $lblReady[$tbp].Size = $DrawSize
        $pnlReady[$tbp].Controls.Add($lblReady[$tbp])

        #endregion lblReady

        }        
}

function btnReadyOnClick($tbp) { ParseInput -str $tbxReady[$tbp].Text -tbp $tbp }
function tbxReadyOnClick($tbp) { $form.acceptButton = $btnReady[$tbp] }


#endregion Ready

#region Matches

$pnlMatches = @{}
$lbxMatches = @{}
$rtbMatches = @{}
$btnResetMatches = @{}
$btnMatchesToggleView = @{}

foreach( $tab in $tabs.keys) {
    foreach( $tbp in $tbpList[$tab]) {

    #region pnlMatches
        
    $pnlMatches[$tbp] = New-Object System.Windows.Forms.Panel
    $pnlMatches[$tbp].Name = "pnlMatches["+$tbp+"]"
    $pnlMatches[$tbp].Visible = $False
    $DrawPoint.X = 0
    $DrawPoint.Y = 0
    $pnlMatches[$tbp].Location = $Drawpoint
    $DrawSize.Width = 690
    $DrawSize.Height = 590
    $pnlMatches[$tbp].Size = $DrawSize
    #$pnlMatches[$tbp].BackColor = "Yellow"
    $tbps[$tbp].Controls.Add($pnlMatches[$tbp])
             
    #endregion pnlMatches

    #region lbxMatches

    $lbxMatches[$tbp] = New-Object System.Windows.Forms.Listbox
    $lbxMatches[$tbp].Name = "lbxMatches["+$tbp+"]"
    $lbxMatches[$tbp].BackColor = $lemon
    $lbxMatches[$tbp].ScrollAlwaysVisible = $True
    #$lbxMatches[$tbp].Visible = $false
    $DrawPoint.X = 20
    $DrawPoint.Y = 10
    $lbxMatches[$tbp].Location = $DrawPoint
    $DrawSize.Width = 550 
    $DrawSize.Height = 470
    $lbxMatches[$tbp].Size = $DrawSize
    $lbxMatches[$tbp].add_SelectedindexChanged({param($sender,$e)   
            lbxMatchesOnIndexChanged -tbp ($sender.name.split('[')[1].split(']')[0])
    })
    drawAsListbox $lbxMatches[$tbp]  
    $pnlMatches[$tbp].Controls.Add($lbxMatches[$tbp])

    #endregion lbxMatches

    #region btnResetMatches

        $btnResetMatches[$tbp] = New-Object System.Windows.Forms.Button
        $btnResetMatches[$tbp].Name = "btnResetMatches["+$tbp+"]"
        $btnResetMatches[$tbp].Text = "Reset"
        $DrawPoint.X = 590
        $DrawPoint.Y = 10
        $btnResetMatches[$tbp].Location = $DrawPoint
        $DrawSize.Width = 80
        $DrawSize.Height = 40
        $btnResetMatches[$tbp].Size = $DrawSize
        $btnResetMatches[$tbp].add_Click({param($sender,$e)
            btnResetMatchesOnClick -tbp ($sender.name.split('[')[1].split(']')[0])
        })
        $pnlMatches[$tbp].Controls.Add($btnResetMatches[$tbp])

    #endregion btnResetMatches
    
    #region btnMatchesToggleView

    $btnMatchesToggleView[$tbp] = New-Object System.Windows.Forms.Button
    $btnMatchesToggleView[$tbp].Name = "btnMatchesToggleView["+$tbp+"]"
    $btnMatchesToggleView[$tbp].Text = ""
    $btnMatchesToggleView[$tbp].Image = $textview
    #$btnMatchesToggleView[$tbp].Visible = $false
    $DrawPoint.X = 590
    $DrawPoint.Y = 70
    $btnMatchesToggleView[$tbp].Location = $DrawPoint
    $DrawSize.Width = 35
    $DrawSize.Height = 35
    $btnMatchesToggleView[$tbp].Size = $DrawSize
    $btnMatchesToggleView[$tbp].add_Click({param($sender,$e)
        btnMatchesToggleViewOnClick -tbp ($sender.name.split('[')[1].split(']')[0])
    })
    $pnlMatches[$tbp].Controls.Add($btnMatchesToggleView[$tbp])

    #endregion btnMatchesToggleView

    #region rtbMatches

    $rtbMatches[$tbp] = New-Object System.Windows.Forms.RichTextBox
    $rtbMatches[$tbp].Name = "rtbMatches["+$tbp+"]"
    $rtbMatches[$tbp].Text = ""
    $rtbMatches[$tbp].Visible = $False
    readOnly $rtbMatches[$tbp]
    $DrawPoint.X = 20
    $DrawPoint.Y = 10
    $rtbMatches[$tbp].Location = $DrawPoint
    $DrawSize.Width = 550 
    $DrawSize.Height = 470
    $rtbMatches[$tbp].Size = $DrawSize
    $pnlMatches[$tbp].Controls.Add($rtbMatches[$tbp])

    #endregion rtbMatches

    }    
}

function btnResetMatchesOnClick($tbp) { enterReady -tbp $tbp }

function lbxMatchesOnIndexChanged ($tbp) { 
    if ($lbxMatches[$tbp].SelectedItem.contains(":") ) { 
        $objClass = "User"
        $obj = getObject $lbxMatches[$tbp].SelectedItem.Split(":")[0] -prop $AdObjProps
        $infoText = getUserInfo $obj
    }
    else { 
        $objClass = "Group"
        $obj = getObject $lbxMatches[$tbp].SelectedItem -prop $AdObjProps
        $infoText =  getGroupInfo $obj
    }
    enterInfo -tbp $tbp -infoText $infoText -obj $obj
}

function btnMatchesToggleViewOnClick ($tbp) {
    toggleView -rtb $rtbMatches[$tbp] -lbx $lbxMatches[$tbp] -btn $btnMatchesToggleView[$tbp] -lbxPic $listview -rtbPic $textview
}

#endregion Matches

#region Info

$pnlInfo = @{}
$rtbInfo = @{}
$btnResetInfo = @{}
$rtbMembership = @{}
$lbxMembership = @{}


foreach( $tab in $tabs.keys) {
    foreach( $tbp in $tbpList[$tab]) {

        #region pnlInfo
        
        $pnlInfo[$tbp] = New-Object System.Windows.Forms.Panel
        $pnlInfo[$tbp].Name = "pnlInfo["+$tbp+"]"
        $pnlInfo[$tbp].Visible = $False
        $DrawPoint.X = 0
        $DrawPoint.Y = 0
        $pnlInfo[$tbp].Location = $Drawpoint
        $DrawSize.Width = 690
        $DrawSize.Height = 150
        $pnlInfo[$tbp].Size = $DrawSize
        #$pnlInfo[$tbp].BackColor = "Orange"
        $tbps[$tbp].Controls.Add($pnlInfo[$tbp])        
             
        #endregion pnlInfo  
        
        #region rtbInfo

        $rtbInfo[$tbp] = New-Object System.Windows.Forms.RichTextBox
        $rtbInfo[$tbp].Name = "rtbInfo["+$tbp+"]"
        $rtbInfo[$tbp].Text = ""
        readOnly $rtbInfo[$tbp]
        $DrawPoint.X = 20
        $DrawPoint.Y = 10
        $rtbInfo[$tbp].Location = $DrawPoint
        $DrawSize.Width = 560 
        $DrawSize.Height = 120
        $rtbInfo[$tbp].Size = $DrawSize
        $pnlInfo[$tbp].Controls.Add($rtbInfo[$tbp])

        #endregion rtbInfo      

        #region btnResetInfo

        $btnResetInfo[$tbp] = New-Object System.Windows.Forms.Button
        $btnResetInfo[$tbp].Name = "btnResetInfo["+$tbp+"]"
        $btnResetInfo[$tbp].Text = "Reset"
        $DrawPoint.X = 590
        $DrawPoint.Y = 10
        $btnResetInfo[$tbp].Location = $DrawPoint
        $DrawSize.Width = 80
        $DrawSize.Height = 40
        $btnResetInfo[$tbp].Size = $DrawSize
        $btnResetInfo[$tbp].add_Click({param($sender,$e)
            btnResetInfoOnClick -tbp ($sender.name.split('[')[1].split(']')[0])
        })
        $pnlInfo[$tbp].Controls.Add($btnResetInfo[$tbp])

        #endregion btnResetInfo

    }
}

function btnResetInfoOnClick($tbp) { enterReady -tbp $tbp }

#endregion Info

#region Ref

$pnlRef = @{}
$lblRefPrompt = @{}
$tbxRef = @{}
$btnRef = @{}

foreach( $tab in $tabs.keys) {
foreach( $tbp in $tbpList[$tab]) {

    #region pnlRef
        
    $pnlRef[$tbp] = New-Object System.Windows.Forms.Panel
    $pnlRef[$tbp].Name = "pnlRef["+$tbp+"]"
    $pnlRef[$tbp].Visible = $False
    $DrawPoint.X = 0
    $DrawPoint.Y = 150
    $pnlRef[$tbp].Location = $Drawpoint
    $DrawSize.Width = 100
    $DrawSize.Height = 350 
    $pnlRef[$tbp].Size = $DrawSize
    #$pnlRef[$tbp].BackColor = $sky
    $tbps[$tbp].Controls.Add($pnlRef[$tbp])
             
    #endregion pnlRef

    #region lblRefPrompt

    $lblRefPrompt[$tbp] = New-Object System.Windows.Forms.Label
    $lblRefPrompt[$tbp].Name = "lblRefPrompt[$tbp]"
    $lblRefPrompt[$tbp].Text = "Enter a ServiceNow Ref to make changes"
    $lblRefPrompt[$tbp].TextAlign = "MiddleCenter" 
    $DrawPoint.X = 10
    $DrawPoint.Y = 50
    $lblRefPrompt[$tbp].Location = $DrawPoint
    $DrawSize.Width = 80
    $DrawSize.Height = 60
    $lblRefPrompt[$tbp].Size = $DrawSize
    $pnlRef[$tbp].Controls.Add($lblRefPrompt[$tbp])

    #endregion lblRefPrompt

    #region tbxRef

    $tbxRef[$tbp] = New-Object System.Windows.Forms.TextBox
    $tbxRef[$tbp].Name = "tbxRef["+$tbp+"]"
    $tbxRef[$tbp].Text = ""
    $DrawPoint.X = 10
    $DrawPoint.Y = 130
    $tbxRef[$tbp].Location = $DrawPoint
    $DrawSize.Width = 80
    $DrawSize.Height = 20
    $tbxRef[$tbp].Size = $DrawSize
    $tbxRef[$tbp].add_TextChanged({param($sender,$e)
        tbxRefOntextChange -tbp ($sender.name.split('[')[1].split(']')[0])
    })
    $pnlRef[$tbp].Controls.Add($tbxRef[$tbp])

    #endregion tbxRef

    #region btnRef

    $btnRef[$tbp] = New-Object System.Windows.Forms.Button
    $btnRef[$tbp].Name = "btnRef["+$tbp+"]"
    $btnRef[$tbp].Text = "Enter Ref"
    $DrawPoint.X = 10
    $DrawPoint.Y = 160
    $btnRef[$tbp].Location = $DrawPoint
    $DrawSize.Width = 80 
    $DrawSize.Height = 40
    $btnRef[$tbp].Size = $DrawSize
    $btnRef[$tbp].add_Click({param($sender,$e)
        btnRefOnClick -tbp ($sender.name.split('[')[1].split(']')[0])
    })
    $pnlRef[$tbp].Controls.Add($btnRef[$tbp])

    #endregion btnRef

    }
}

Function btnRefOnClick($tbp) { enterSetActions -tbp $tbp }
Function tbxRefOnTextChange($tbp) { $form.AcceptButton = $btnRef[$tbp] }

#endregion Ref

#region SetActions

$pnlSetActions = @{}
$lblRefSet = @{}
$btnAddgroups = @{}
$btnRemGroups = @{}
$btnAddRemGroups = @{}
$btnClearGroups = @{}
$btnCloneGroups = @{}

foreach( $tab in $tabs.keys) {
foreach( $tbp in $tbpList[$tab]) {

    #region pnlSetActions
        
    $pnlSetActions[$tbp] = New-Object System.Windows.Forms.Panel
    $pnlSetActions[$tbp].Name = "pnlSetActions["+$tbp+"]"
    $pnlSetActions[$tbp].Visible = $False
    $DrawPoint.X = 0
    $DrawPoint.Y = 150
    $pnlSetActions[$tbp].Location = $Drawpoint
    $DrawSize.Width = 100
    $DrawSize.Height = 350
    $pnlSetActions[$tbp].Size = $DrawSize
    #$pnlSetActions[$tbp].BackColor = $grey
    $tbps[$tbp].Controls.Add($pnlSetActions[$tbp])
             
    #endregion pnlSetActions

    #region lblRefSet

    $lblRefSet[$tbp] = New-Object System.Windows.Forms.Label
    $lblRefSet[$tbp].Name = "lblRefSet[$tbp]"
    $lblRefSet[$tbp].Text = "SNOW REF"
    $lblRefSet[$tbp].TextAlign = "MiddleCenter" 
    $DrawPoint.X = 10
    $DrawPoint.Y = 0
    $lblRefSet[$tbp].Location = $DrawPoint
    $DrawSize.Width = 80 
    $DrawSize.Height = 20
    $lblRefSet[$tbp].Size = $DrawSize
    $pnlSetActions[$tbp].Controls.Add($lblRefSet[$tbp])

    #endregion lblRefSet

    #region btnAddGroups

    $btnAddGroups[$tbp] = New-Object System.Windows.Forms.Button
    $btnAddGroups[$tbp].Name = "btnAddGroups["+$tbp+"]"
    $btnAddGroups[$tbp].Text = "Add Groups"
    $btnAddGroups[$tbp].BackColor = $warnOrange
    $DrawPoint.X = 10
    $DrawPoint.Y = 20
    $btnAddGroups[$tbp].Location = $DrawPoint
    $DrawSize.Width = 80
    $DrawSize.Height = 50
    $btnAddGroups[$tbp].Size = $DrawSize
    $btnAddGroups[$tbp].add_Click({param($sender,$e)
        btnAddGroupsOnClick -tbp ($sender.name.split('[')[1].split(']')[0])
    })
    $pnlSetActions[$tbp].Controls.Add($btnAddGroups[$tbp])

    #endregion btnAddGroups

    #region btnRemGroups

    $btnRemGroups[$tbp] = New-Object System.Windows.Forms.Button
    $btnRemGroups[$tbp].Name = "btnRemGroups["+$tbp+"]"
    $btnRemGroups[$tbp].Text = "Remove Groups"
    $btnRemGroups[$tbp].BackColor = $warnPink
    $DrawPoint.X = 10
    $DrawPoint.Y = 90
    $btnRemGroups[$tbp].Location = $DrawPoint
    $DrawSize.Width = 80
    $DrawSize.Height = 50
    $btnRemGroups[$tbp].Size = $DrawSize
    $btnRemGroups[$tbp].add_Click({param($sender,$e)
        btnRemGroupsOnClick -tbp ($sender.name.split('[')[1].split(']')[0])
    })
    $pnlSetActions[$tbp].Controls.Add($btnRemGroups[$tbp])

    #endregion btnRemoveGroups

    #region btnAddRemGroups

    $btnAddRemGroups[$tbp] = New-Object System.Windows.Forms.Button
    $btnAddRemGroups[$tbp].Name = "btnAddRemGroups["+$tbp+"]"
    $btnAddRemGroups[$tbp].Text = "Add + Remove Groups"
    $btnAddRemGroups[$tbp].BackColor = $warnOrange
    $DrawPoint.X = 10
    $DrawPoint.Y = 160
    $btnAddRemGroups[$tbp].Location = $DrawPoint
    $DrawSize.Width = 80
    $DrawSize.Height = 50
    $btnAddRemGroups[$tbp].Size = $DrawSize
    $btnAddRemGroups[$tbp].add_Click({param($sender,$e)
        btnAddRemGroupsOnClick -tbp ($sender.name.split('[')[1].split(']')[0])
    })
    $pnlSetActions[$tbp].Controls.Add($btnAddRemGroups[$tbp])

    #endregion btnAddRemGroups

    #region btnClearGroups

    $btnClearGroups[$tbp] = New-Object System.Windows.Forms.Button
    $btnClearGroups[$tbp].Name = "btnClearGroups["+$tbp+"]"
    $btnClearGroups[$tbp].Text = "Clear Groups"
    $btnClearGroups[$tbp].BackColor = $warnRed
    $DrawPoint.X = 10
    $DrawPoint.Y = 230
    $btnClearGroups[$tbp].Location = $DrawPoint
    $DrawSize.Width = 80 
    $DrawSize.Height = 50
    $btnClearGroups[$tbp].Size = $DrawSize
    $btnClearGroups[$tbp].add_Click({param($sender,$e)
        btnClearGroupsOnClick -tbp ($sender.name.split('[')[1].split(']')[0])
    })
    $pnlSetActions[$tbp].Controls.Add($btnClearGroups[$tbp])

    #endregion btnClearGroups

    #region btnCloneGroups

    $btnCloneGroups[$tbp] = New-Object System.Windows.Forms.Button
    $btnCloneGroups[$tbp].Name = "btnCloneGroups["+$tbp+"]"
    $btnCloneGroups[$tbp].Text = "Clone Groups"
    $btnCloneGroups[$tbp].BackColor = $warnRed
    $btnCloneGroups[$tbp].Visible = $false
    $DrawPoint.X = 10
    $DrawPoint.Y = 300
    $btnCloneGroups[$tbp].Location = $DrawPoint
    $DrawSize.Width = 80 
    $DrawSize.Height = 50
    $btnCloneGroups[$tbp].Size = $DrawSize
    $btnCloneGroups[$tbp].add_Click({param($sender,$e)
        btnCloneGroupsOnClick -tbp ($sender.name.split('[')[1].split(']')[0])
    })
    $pnlSetActions[$tbp].Controls.Add($btnCloneGroups[$tbp])

    #endregion btnCloneGroups

    }
}

Function btnAddGroupsOnClick( $tbp ) {
    enterEnact -tbp $tbp  -action "Add"
}

Function btnRemGroupsOnClick( $tbp ) {
    enterEnact -tbp $tbp  -action "Remove"
}

Function btnAddRemGroupsOnClick( $tbp ) {
    enterEnact -tbp $tbp  -action "Add + Remove" 
}

Function btnClearGroupsOnClick( $tbp ) {
    enterEnact -tbp $tbp  -action "Remove" -remgroups $rtbMembership[$tbp].text
}


#endregion SetActions

#region Membership

$pnlMembership = @{}
$lblMembership = @{} 
$lblRemGroups = @{}
$lblAddGroups = @{}
$rtbMembership = @{}
$rtbMidThird = @{}
$rtbRightThird = @{}
$rtbRightHalf = @{} 

foreach( $tab in $tabs.keys) {
    foreach( $tbp in $tbpList[$tab]) {

    #region pnlMembership
        
    $pnlMembership[$tbp] = New-Object System.Windows.Forms.Panel
    $pnlMembership[$tbp].Name = "pnlMembership["+$tbp+"]"
    $pnlMembership[$tbp].Visible = $False
    $DrawPoint.X = 100
    $DrawPoint.Y = 150
    $pnlMembership[$tbp].Location = $Drawpoint
    $DrawSize.Width = 480 
    $DrawSize.Height = 350
    $pnlMembership[$tbp].Size = $DrawSize
    #$pnlMembership[$tbp].BackColor = $lemon
    $tbps[$tbp].Controls.Add($pnlMembership[$tbp])
             
    #endregion pnlMembership

    #region rtbMembership

    $rtbMembership[$tbp] = New-Object System.Windows.Forms.RichTextBox
    $rtbMembership[$tbp].Name = "rtbMembership["+$tbp+"]"
    $rtbMembership[$tbp].Text = ""
    #hide $rtbMembership[$tbp]
    readOnly $rtbMembership[$tbp]
    $DrawPoint.X = 0
    $DrawPoint.Y = 20
    $rtbMembership[$tbp].Location = $DrawPoint
    $DrawSize.Width = 480
    $DrawSize.Height = 330
    $rtbMembership[$tbp].Size = $DrawSize
    $pnlMembership[$tbp].Controls.Add($rtbMembership[$tbp])

    #endregion rtbMembership

    #region lblMembership

    $lblMembership[$tbp] = New-Object System.Windows.Forms.Label
    $lblMembership[$tbp].Name = "lblMembership[$tbp]"
    $lblMembership[$tbp].Text = "Groups"
    $lblMembership[$tbp].TextAlign = "MiddleLeft" 
    $DrawPoint.X = 0
    $DrawPoint.Y = 0
    $lblMembership[$tbp].Location = $DrawPoint
    $DrawSize.Width = 480
    $DrawSize.Height = 20
    $lblMembership[$tbp].Size = $DrawSize
    $pnlMembership[$tbp].Controls.Add($lblMembership[$tbp])

    #endregion lblMembership

    #region lblRemGroups

    $lblRemGroups[$tbp] = New-Object System.Windows.Forms.Label
    $lblRemGroups[$tbp].Name = "lblRemGroups[$tbp]"
    $lblRemGroups[$tbp].Text = "Groups to Remove"
    $lblRemGroups[$tbp].TextAlign = "MiddleCenter" 
    hide $lblRemGroups[$tbp]
    $DrawPoint.X = 160
    $DrawPoint.Y = 0
    $lblRemGroups[$tbp].Location = $DrawPoint
    $DrawSize.Width = 160
    $DrawSize.Height = 20
    $lblRemGroups[$tbp].Size = $DrawSize
    $pnlMembership[$tbp].Controls.Add($lblRemGroups[$tbp])

    #endregion lblRemGroups

    #region lblAddGroups

    $lblAddGroups[$tbp] = New-Object System.Windows.Forms.Label
    $lblAddGroups[$tbp].Name = "lblAddGroups[$tbp]"
    $lblAddGroups[$tbp].Text = "Groups to Add"
    $lblAddGroups[$tbp].TextAlign = "MiddleRight" 
    hide $lblAddGroups[$tbp]
    $DrawPoint.X = 320
    $DrawPoint.Y = 0
    $lblAddGroups[$tbp].Location = $DrawPoint
    $DrawSize.Width = 160
    $DrawSize.Height = 20
    $lblAddGroups[$tbp].Size = $DrawSize
    $pnlMembership[$tbp].Controls.Add($lblAddGroups[$tbp])

    #endregion lblAddGroups
    
    #region rtbMidThird

    $rtbMidThird[$tbp] = New-Object System.Windows.Forms.RichTextBox
    $rtbMidThird[$tbp].Name = "rtbMidThird["+$tbp+"]"
    $rtbMidThird[$tbp].Text = "mid"
    $rtbMidThird[$tbp].Visible = $False
    $DrawPoint.X = 160
    $DrawPoint.Y = 20
    $rtbMidThird[$tbp].Location = $DrawPoint
    $DrawSize.Width = 160 
    $DrawSize.Height = 330
    $rtbMidThird[$tbp].Size = $DrawSize
    $pnlMembership[$tbp].Controls.Add($rtbMidThird[$tbp])

    #endregion rtbMidThird

    #region rtbRightThird

    $rtbRightThird[$tbp] = New-Object System.Windows.Forms.RichTextBox
    $rtbRightThird[$tbp].Name = "rtbRightThird["+$tbp+"]"
    $rtbRightThird[$tbp].Text = "right third"
    $rtbRightThird[$tbp].Visible = $False
    $DrawPoint.X = 320
    $DrawPoint.Y = 20
    $rtbRightThird[$tbp].Location = $DrawPoint
    $DrawSize.Width = 160 
    $DrawSize.Height = 330
    $rtbRightThird[$tbp].Size = $DrawSize
    $pnlMembership[$tbp].Controls.Add($rtbRightThird[$tbp])

    #endregion rtbRightThird
    
    #region rtbRightHalf

    $rtbRightHalf[$tbp] = New-Object System.Windows.Forms.RichTextBox
    $rtbRightHalf[$tbp].Name = "rtbRightHalf["+$tbp+"]"
    $rtbRightHalf[$tbp].Text = "Right Half"
    $rtbRightHalf[$tbp].Visible = $False
    $DrawPoint.X = 240
    $DrawPoint.Y = 20
    $rtbRightHalf[$tbp].Location = $DrawPoint
    $DrawSize.Width = 240 
    $DrawSize.Height = 330
    $rtbRightHalf[$tbp].Size = $DrawSize
    $pnlMembership[$tbp].Controls.Add($rtbRightHalf[$tbp])

    #endregion rtbRightHalf

    #region lbxMembership

    $lbxMembership[$tbp] = New-Object System.Windows.Forms.Listbox
    $lbxMembership[$tbp].Name = "lbxMembership["+$tbp+"]"
    $lbxMembership[$tbp].Visible = $false
    drawAsListbox $lbxMembership[$tbp]
    $lbxMembership[$tbp].BackColor = $lemon
    $lbxMembership[$tbp].ScrollAlwaysVisible = $True
    $DrawPoint.X = 0
    $DrawPoint.Y = 20
    $lbxMembership[$tbp].Location = $DrawPoint
    $DrawSize.Width = 480
    $DrawSize.Height = 330
    $lbxMembership[$tbp].Size = $DrawSize
    $lbxMembership[$tbp].add_SelectedindexChanged({param($sender,$e)   
        lbxMembershipOnIndexChanged -tbp ($sender.name.split('[')[1].split(']')[0])
    })
    $pnlMembership[$tbp].Controls.Add($lbxMembership[$tbp])

    #endregion lbxMembership
    }    
}

Function lbxMembershipOnIndexChanged($tbp) {
    
    $obj = getObject $lbxMembership[$tbp].SelectedItem -prop $AdObjProps
    $infoText =  getObjInfo $obj
        
    #toggleView -lbx $lbxMembership[$tbp] -rtb $rtbMembership[$tbp] -btn $btnToggleMembershipView[$tbp] -lbxPic $listview -rtbPic $textview  
    enterReady -tbp $tbp
    exitReady -tbp $tbp
      
    enterInfo -tbp $tbp -infoText $infoText -obj $obj    
}

#endregion Membership

#region Enact

$pnlEnact = @{}
$btnCancelEnact = @{}
$btnEnact = @{}
$btnToggleMembershipView = @{}
$btnToggleMembershipDirection = @{}
$btnELV = @{}

foreach( $tab in $tabs.keys) {
foreach( $tbp in $tbpList[$tab]) {

    #region pnlEnact
        
    $pnlEnact[$tbp] = New-Object System.Windows.Forms.Panel
    $pnlEnact[$tbp].Name = "pnlEnact["+$tbp+"]"
    $pnlEnact[$tbp].Visible = $False
    $DrawPoint.X = 580
    $DrawPoint.Y = 150
    $pnlEnact[$tbp].Location = $Drawpoint
    $DrawSize.Width = 110
    $DrawSize.Height = 350 
    $pnlEnact[$tbp].Size = $DrawSize
    #$pnlEnact[$tbp].BackColor = $grey
    $tbps[$tbp].Controls.Add($pnlEnact[$tbp])
             
    #endregion pnlEnact

    #region btnCancelEnact

    $btnCancelEnact[$tbp] = New-Object System.Windows.Forms.Button
    $btnCancelEnact[$tbp].Name = "btnCancelEnact["+$tbp+"]"
    $btnCancelEnact[$tbp].Text = "Cancel"
    $btnCancelEnact[$tbp].Visible = $false
    $DrawPoint.X = 10
    $DrawPoint.Y = 20
    $btnCancelEnact[$tbp].Location = $DrawPoint
    $DrawSize.Width = 80
    $DrawSize.Height = 40
    $btnCancelEnact[$tbp].Size = $DrawSize
    $btnCancelEnact[$tbp].add_Click({param($sender,$e)
        btnCancelEnactOnClick -tbp ($sender.name.split('[')[1].split(']')[0])
    })
    $pnlEnact[$tbp].Controls.Add($btnCancelEnact[$tbp])

    #endregion btnCancelEnact

    #region btnEnact

    $btnEnact[$tbp] = New-Object System.Windows.Forms.Button
    $btnEnact[$tbp].Name = "btnEnact["+$tbp+"]"
    $btnEnact[$tbp].Text = "Enact"
    $btnEnact[$tbp].Visible = $false
    $DrawPoint.X = 10
    $DrawPoint.Y = 90
    $btnEnact[$tbp].Location = $DrawPoint
    $DrawSize.Width = 80
    $DrawSize.Height = 40
    $btnEnact[$tbp].Size = $DrawSize
    $btnEnact[$tbp].add_Click({param($sender,$e)
        btnEnactOnClick -tbp ($sender.name.split('[')[1].split(']')[0])
    })
    $pnlEnact[$tbp].Controls.Add($btnEnact[$tbp])

    #endregion btnEnact

    #region btnToggleMembershipView

    $btnToggleMembershipView[$tbp] = New-Object System.Windows.Forms.Button
    $btnToggleMembershipView[$tbp].Name = "btnToggleMembershipView["+$tbp+"]"
    $btnToggleMembershipView[$tbp].Text = ""
    $btnToggleMembershipView[$tbp].Image = $listview
    $DrawPoint.X = 10
    $DrawPoint.Y = 160
    $btnToggleMembershipView[$tbp].Location = $DrawPoint
    $DrawSize.Width = 35
    $DrawSize.Height = 35
    $btnToggleMembershipView[$tbp].Size = $DrawSize
    $btnToggleMembershipView[$tbp].add_Click({param($sender,$e)
        btnToggleMembershipViewOnClick -tbp ($sender.name.split('[')[1].split(']')[0])
    })
    $pnlEnact[$tbp].Controls.Add($btnToggleMembershipView[$tbp])

    #endregion btnToggleMembershipView

    #region btnToggleMembershipDirection

    $btnToggleMembershipDirection[$tbp] = New-Object System.Windows.Forms.Button
    $btnToggleMembershipDirection[$tbp].Name = "btnToggleMembershipDirection["+$tbp+"]"
    $btnToggleMembershipDirection[$tbp].Text = ""
    $btnToggleMembershipDirection[$tbp].Image = $memberOfPic
    $btnToggleMembershipDirection[$tbp].Visible = $false
    $DrawPoint.X = 55
    $DrawPoint.Y = 160
    $btnToggleMembershipDirection[$tbp].Location = $DrawPoint
    $DrawSize.Width = 35
    $DrawSize.Height = 35
    $btnToggleMembershipDirection[$tbp].Size = $DrawSize
    $btnToggleMembershipDirection[$tbp].add_Click({param($sender,$e)
        btnToggleMembershipDirectionOnClick -tbp ($sender.name.split('[')[1].split(']')[0])
    })
    $pnlEnact[$tbp].Controls.Add($btnToggleMembershipDirection[$tbp])

    #endregion btnToggleMembershipDirection

    #region btnELV

    $btnELV[$tbp] = New-Object System.Windows.Forms.Button
    $btnELV[$tbp].Name = "btnELV["+$tbp+"]"
    $btnELV[$tbp].Text = "ELV"
    $btnELV[$tbp].Visible = $false
    $DrawPoint.X = 55
    $DrawPoint.Y = 160
    $btnELV[$tbp].Location = $DrawPoint
    $DrawSize.Width = 35
    $DrawSize.Height = 35
    $btnELV[$tbp].Size = $DrawSize
    $btnELV[$tbp].add_Click({param($sender,$e)
        btnELVOnClick -tbp ($sender.name.split('[')[1].split(']')[0])
    })
    $pnlEnact[$tbp].Controls.Add($btnELV[$tbp])

    #endregion btnELV

    }
}

Function btnCancelEnactOnClick($tbp) { exitEnact -tbp $tbp }
Function btnToggleMembershipViewOnClick($tbp) { toggleView -lbx $lbxMembership[$tbp] -rtb $rtbMembership[$tbp] -btn $btnToggleMembershipView[$tbp] -lbxPic $listview -rtbPic $textview }
Function btnToggleMembershipDirectionOnClick($tbp) { toggleMembershipDirection -tbp $tbp }
Function btnEnactOnClick($tbp) { enact -tbp $tbp }
Function btnELVOnClick($tbp) { enterELV -tbp $tbp }

#endregion Enact

#region GroupResults

$pnlGroupResults = @{}
$rtbGroupResults = @{}
$btnCloseGroupResults = @{}

foreach( $tab in $tabs.keys) {
foreach( $tbp in $tbpList[$tab]) {

    #region pnlGroupResults
        
    $pnlGroupResults[$tbp] = New-Object System.Windows.Forms.Panel
    $pnlGroupResults[$tbp].Name = "pnlGroupResults["+$tbp+"]"
    $pnlGroupResults[$tbp].Visible = $False
    $DrawPoint.X = 0
    $DrawPoint.Y = 500
    $pnlGroupResults[$tbp].Location = $Drawpoint
    $DrawSize.Width = 690
    $DrawSize.Height = 290
    $pnlGroupResults[$tbp].Size = $DrawSize
    #$pnlGroupResults[$tbp].BackColor = "Cyan"
    $tbps[$tbp].Controls.Add($pnlGroupResults[$tbp])
             
    #endregion pnlGroupResults

    #region btnCloseGroupResults

    $btnCloseGroupResults[$tbp] = New-Object System.Windows.Forms.Button
    $btnCloseGroupResults[$tbp].Name = "btnCloseGroupResults["+$tbp+"]"
    $btnCloseGroupResults[$tbp].Text = "Done"
    $DrawPoint.X = 600
    $DrawPoint.Y = 120
    $btnCloseGroupResults[$tbp].Location = $DrawPoint
    $DrawSize.Width = 80
    $DrawSize.Height = 40
    $btnCloseGroupResults[$tbp].Size = $DrawSize
    $btnCloseGroupResults[$tbp].add_Click({param($sender,$e)
        btnCloseGroupResultsOnClick -tbp ($sender.name.split('[')[1].split(']')[0])
    })
    $pnlGroupResults[$tbp].Controls.Add($btnCloseGroupResults[$tbp])

    #endregion btnCloseGroupResults

    #region rtbGroupResults

    $rtbGroupResults[$tbp] = New-Object System.Windows.Forms.RichTextBox
    $rtbGroupResults[$tbp].Name = "rtbGroupResults["+$tbp+"]"
    $rtbGroupResults[$tbp].Text = ""
    $DrawPoint.X = 0
    $DrawPoint.Y = 0
    $rtbGroupResults[$tbp].Location = $DrawPoint
    $DrawSize.Width = 590
    $DrawSize.Height = 290
    $rtbGroupResults[$tbp].Size = $DrawSize
    $pnlGroupResults[$tbp].Controls.Add($rtbGroupResults[$tbp])

    #endregion rtbGroupResults

    }
}

Function btnCloseGroupResultsOnClick($tbp) {   
    exitGroupResults -tbp $tbp
}

#endregion GroupResults

#region Compare

$pnlCompare = @{}
$rtbLeftOnly = @{}
$rtbRightOnly = @{}
$rtbBothLeftandRight = @{}
$rtbEitherLeftOrRight = @{}
$lblLeftOnly = @{}
$lblRightOnly = @{}
$lblBothLeftandRight = @{}
$lblEitherLeftOrRight = @{}
$btnCloseLeftCompare = @{}
$btnCloseRightCompare = @{}


foreach( $tab in $tabs.keys) {
foreach( $tbp in $tbpList[$tab]) {

    #region pnlCompare
        
    $pnlCompare[$tbp] = New-Object System.Windows.Forms.Panel
    $pnlCompare[$tbp].Name = "pnlCompare["+$tbp+"]"
    $pnlCompare[$tbp].Visible = $False
    $DrawPoint.X = 0
    $DrawPoint.Y = 500
    $pnlCompare[$tbp].Location = $Drawpoint
    $DrawSize.Width = 690
    $DrawSize.Height = 290
    $pnlCompare[$tbp].Size = $DrawSize
    #$pnlCompare[$tbp].BackColor = "Cyan"
    $tbps[$tbp].Controls.Add($pnlCompare[$tbp])
             
    #endregion pnlCompare

    #region left

    if($tab -eq "Main") {

        #region rtbLeftOnly

        $rtbLeftOnly[$tbp] = New-Object System.Windows.Forms.RichTextBox
        $rtbLeftOnly[$tbp].Name = "rtbLeftOnly["+$tbp+"]"
        $rtbLeftOnly[$tbp].Text = ""
        $DrawPoint.X = 0
        $DrawPoint.Y = 20
        $rtbLeftOnly[$tbp].Location = $DrawPoint
        $DrawSize.Width = 340
        $DrawSize.Height = 250
        $rtbLeftOnly[$tbp].Size = $DrawSize
        $pnlCompare[$tbp].Controls.Add($rtbLeftOnly[$tbp])

        #endregion rtbLeftOnly 
        
        #region rtbBothLeftAndRight   

        $rtbBothLeftAndRight[$tbp] = New-Object System.Windows.Forms.RichTextBox
        $rtbBothLeftAndRight[$tbp].Name = "rtbBothLeftAndRight["+$tbp+"]"
        $rtbBothLeftAndRight[$tbp].Text = ""
        $DrawPoint.X = 340
        $DrawPoint.Y = 20
        $rtbBothLeftAndRight[$tbp].Location = $DrawPoint
        $DrawSize.Width = 340 
        $DrawSize.Height = 250
        $rtbBothLeftAndRight[$tbp].Size = $DrawSize
        $pnlCompare[$tbp].Controls.Add($rtbBothLeftAndRight[$tbp])

        #endregion rtbBothLeftAndRight  
        
        #region lblLeftOnly

        $lblLeftOnly[$tbp] = New-Object System.Windows.Forms.Label
        $lblLeftOnly[$tbp].Name = "lblLeftOnly[$tbp]"
        $lblLeftOnly[$tbp].Text = "Only"
        $lblLeftOnly[$tbp].TextAlign = "MiddleCenter" 
        $DrawPoint.X = 70
        $DrawPoint.Y = 0
        $lblLeftOnly[$tbp].Location = $DrawPoint
        $DrawSize.Width = 200
        $DrawSize.Height = 20
        $lblLeftOnly[$tbp].Size = $DrawSize
        $pnlCompare[$tbp].Controls.Add($lblLeftOnly[$tbp])

        #endregion lblLeftOnly  

        #region lblBothLeftAndRight

        $lblBothLeftAndRight[$tbp] = New-Object System.Windows.Forms.Label
        $lblBothLeftAndRight[$tbp].Name = "lblBothLeftAndRight[$tbp]"
        $lblBothLeftAndRight[$tbp].Text = ""
        $lblBothLeftAndRight[$tbp].TextAlign = "MiddleCenter" 
        $DrawPoint.X = 410
        $DrawPoint.Y = 0
        $lblBothLeftAndRight[$tbp].Location = $DrawPoint
        $DrawSize.Width = 200
        $DrawSize.Height = 20
        $lblBothLeftAndRight[$tbp].Size = $DrawSize
        $pnlCompare[$tbp].Controls.Add($lblBothLeftAndRight[$tbp])

        #endregion lblBothLeftAndRight  

        #region btnCloseLeftCompare

        $btnCloseLeftCompare[$tbp] = New-Object System.Windows.Forms.Button
        $btnCloseLeftCompare[$tbp].Name = "btnCloseLeftCompare["+$tbp+"]"
        $btnCloseLeftCompare[$tbp].Text = "X"
        $DrawPoint.X = 620
        $DrawPoint.Y = 0
        $btnCloseLeftCompare[$tbp].Location = $DrawPoint
        $DrawSize.Width = 60
        $DrawSize.Height = 20
        $btnCloseLeftCompare[$tbp].Size = $DrawSize
        $btnCloseLeftCompare[$tbp].add_Click({btnCloseCompares})
        $pnlCompare[$tbp].Controls.Add($btnCloseLeftCompare[$tbp])

        #endregion btnCloseLeftCompare

    }

    #endregion left

    #region right
    

    if($tab -eq "Aux") {

        #region rtbEitherLeftOrRight

        $rtbEitherLeftOrRight[$tbp] = New-Object System.Windows.Forms.RichTextBox
        $rtbEitherLeftOrRight[$tbp].Name = "rtbEitherLeftOrRight["+$tbp+"]"
        $rtbEitherLeftOrRight[$tbp].Text = ""
        $DrawPoint.X = 0
        $DrawPoint.Y = 20
        $rtbEitherLeftOrRight[$tbp].Location = $DrawPoint
        $DrawSize.Width = 340
        $DrawSize.Height = 250
        $rtbEitherLeftOrRight[$tbp].Size = $DrawSize
        $pnlCompare[$tbp].Controls.Add($rtbEitherLeftOrRight[$tbp])

        #endregion rtbEitherLeftOrRight    

        #region rtbRightOnly

        $rtbRightOnly[$tbp] = New-Object System.Windows.Forms.RichTextBox
        $rtbRightOnly[$tbp].Name = "rtbRightOnly["+$tbp+"]"
        $rtbRightOnly[$tbp].Text = ""
        $DrawPoint.X = 340
        $DrawPoint.Y = 20
        $rtbRightOnly[$tbp].Location = $DrawPoint
        $DrawSize.Width = 340
        $DrawSize.Height = 250
        $rtbRightOnly[$tbp].Size = $DrawSize
        $pnlCompare[$tbp].Controls.Add($rtbRightOnly[$tbp])

        #endregion rtbRightOnly  
        
        #region lblEitherLeftOrRight

        $lblEitherLeftOrRight[$tbp] = New-Object System.Windows.Forms.Label
        $lblEitherLeftOrRight[$tbp].Name = "lblEitherLeftOrRight[$tbp]"
        $lblEitherLeftOrRight[$tbp].Text = "spacing text just to test with"
        $lblEitherLeftOrRight[$tbp].TextAlign = "MiddleCenter" 
        $DrawPoint.X = 70
        $DrawPoint.Y = 0
        $lblEitherLeftOrRight[$tbp].Location = $DrawPoint
        $DrawSize.Width = 200
        $DrawSize.Height = 20
        $lblEitherLeftOrRight[$tbp].Size = $DrawSize
        $pnlCompare[$tbp].Controls.Add($lblEitherLeftOrRight[$tbp])

        #endregion lblEitherLeftOrRight  

        #region lblRightOnly

        $lblRightOnly[$tbp] = New-Object System.Windows.Forms.Label
        $lblRightOnly[$tbp].Name = "lblRightOnly[$tbp]"
        $lblRightOnly[$tbp].Text = "spacing text just to test with"
        $lblRightOnly[$tbp].TextAlign = "MiddleCenter" 
        $DrawPoint.X = 410
        $DrawPoint.Y = 0
        $lblRightOnly[$tbp].Location = $DrawPoint
        $DrawSize.Width = 200
        $DrawSize.Height = 20
        $lblRightOnly[$tbp].Size = $DrawSize
        $pnlCompare[$tbp].Controls.Add($lblRightOnly[$tbp])

        #endregion lblRightOnly  

        #region btnCloseRightCompare

        $btnCloseRightCompare[$tbp] = New-Object System.Windows.Forms.Button
        $btnCloseRightCompare[$tbp].Name = "btnCloseRightCompare["+$tbp+"]"
        $btnCloseRightCompare[$tbp].Text = "X"
        $DrawPoint.X = 620
        $DrawPoint.Y = 0
        $btnCloseRightCompare[$tbp].Location = $DrawPoint
        $DrawSize.Width = 60
        $DrawSize.Height = 20
        $btnCloseRightCompare[$tbp].Size = $DrawSize
        $btnCloseRightCompare[$tbp].add_Click({btnCloseCompares})
        $pnlCompare[$tbp].Controls.Add($btnCloseRightCompare[$tbp])

        #endregion btnCloseRightCompare

    }

    #>
    #endregion right

    }
}

Function btnCloseCompares { exitCompare }


#endregion Compare

#region ELV
#endregion ELV

#region MovePDUsers

    #region btnMoveMove

    $btnMoveMove = New-Object System.Windows.Forms.Button
    $btnMoveMove.Name = "btnMoveMove"
    $btnMoveMove.Text = "Move"
    $DrawPoint.X = 560
    $DrawPoint.Y = 10
    $btnMoveMove.Location = $DrawPoint
    $DrawSize.Width = 80
    $DrawSize.Height = 40
    $btnMoveMove.Size = $DrawSize
    $btnMoveMove.add_Click({btnMoveMoveOnClick})
    $tbpMoveUsers.Controls.Add($btnMoveMove)

    Function btnMoveMoveOnClick { 
        $btnMoveMove.Visible = $False
        #verifyDataDate
        MovePdUsers -correctPD $tbxMoveCorrectPd.Text -users (RTB2List -rtb $rtbMoveUsers) -copyGroups (RTB2List -rtb $rtbCopyAccess) -removeGroups (RTB2List -rtb $rtbRemovePDs) -rtbOut $rtbMoveResults        
    }


    #endregion btnMoveMove

    #region lblMoveCorrectPd

    $lblMoveCorrectPd = New-Object System.Windows.Forms.Label
    $lblMoveCorrectPd.Name = "lblMoveCorrectPd"
    $lblMoveCorrectPd.Text = "Correct PD:"
    $lblMoveCorrectPd.TextAlign = "MiddleCenter" 
    $DrawPoint.X = 0
    $DrawPoint.Y = 10
    $lblMoveCorrectPd.Location = $DrawPoint
    $DrawSize.Width = 80
    $DrawSize.Height = 20
    $lblMoveCorrectPd.Size = $DrawSize
    $tbpMoveUsers.Controls.Add($lblMoveCorrectPd)

    #endregion lblMoveCorrectPd

    #region tbxMoveCorrectPd

    $tbxMoveCorrectPd = New-Object System.Windows.Forms.TextBox
    $tbxMoveCorrectPd.Name = "tbxMoveCorrectPd"
    $tbxMoveCorrectPd.Text = ""
    $DrawPoint.X = 80
    $DrawPoint.Y = 10
    $tbxMoveCorrectPd.Location = $DrawPoint
    $DrawSize.Width = 300
    $DrawSize.Height = 20 
    $tbxMoveCorrectPd.Size = $DrawSize
    $tbpMoveUsers.Controls.Add($tbxMoveCorrectPd)

    #endregion tbxMoveCorrectPd

    #region btnMoveClear

    $btnMoveClear = New-Object System.Windows.Forms.Button
    $btnMoveClear.Name = "btnMoveClear"
    $btnMoveClear.Text = "Clear"
    $DrawPoint.X = 560
    $DrawPoint.Y = 250
    $btnMoveClear.Location = $DrawPoint
    $DrawSize.Width = 80
    $DrawSize.Height = 40
    $btnMoveClear.Size = $DrawSize
    $btnMoveClear.add_Click({btnMoveClearOnClick})
    $tbpMoveUsers.Controls.Add($btnMoveClear)

    Function btnMoveClearOnClick{
        $tbxMoveCorrectPd.Text = ""
        $rtbMoveUsers.Text = ""
        $rtbRemovePDs.Text = ""
        $rtbCopyAccess.Text = ""
        $rtbMoveResults.Text = ""
        $btnMoveMove.Visible = $True
    }

    #endregion btnMoveClear

    #region rtbMoveUsers

    $rtbMoveUsers = New-Object System.Windows.Forms.RichTextBox
    $rtbMoveUsers.Name = "rtbMoveUsers"
    $rtbMoveUsers.Text = ""
    $DrawPoint.X = 10
    $DrawPoint.Y = 80
    $rtbMoveUsers.Location = $DrawPoint
    $DrawSize.Width = 80 
    $DrawSize.Height = 150
    $rtbMoveUsers.Size = $DrawSize
    $tbpMoveUsers.Controls.Add($rtbMoveUsers)

    #endregion rtbMoveUsers

    #region lblMoveUsers

    $lblMoveUsers = New-Object System.Windows.Forms.Label
    $lblMoveUsers.Name = "lblMoveUsers"
    $lblMoveUsers.Text = "Users"
    $lblMoveUsers.TextAlign = "BottomCenter" 
    $DrawPoint.X = 10
    $DrawPoint.Y = 60
    $lblMoveUsers.Location = $DrawPoint
    $DrawSize.Width = 80 
    $DrawSize.Height = 20
    $lblMoveUsers.Size = $DrawSize
    $tbpMoveUsers.Controls.Add($lblMoveUsers)

    #endregion lblMoveUsers

    #region rtbRemovePDs

    $rtbRemovePDs = New-Object System.Windows.Forms.RichTextBox
    $rtbRemovePDs.Name = "rtbRemovePDs"
    $rtbRemovePDs.Text = ""
    $DrawPoint.X = 100
    $DrawPoint.Y = 80
    $rtbRemovePDs.Location = $DrawPoint
    $DrawSize.Width = 280 
    $DrawSize.Height = 150
    $rtbRemovePDs.Size = $DrawSize
    $tbpMoveUsers.Controls.Add($rtbRemovePDs)

    #endregion rtbRemovePDs

    #region lblRemovePDs

    $lblRemovePDs = New-Object System.Windows.Forms.Label
    $lblRemovePDs.Name = "lblRemovePDs"
    $lblRemovePDs.Text = "PDs to be removed"
    $lblRemovePDs.TextAlign = "BottomCenter" 
    $DrawPoint.X = 100
    $DrawPoint.Y = 60
    $lblRemovePDs.Location = $DrawPoint
    $DrawSize.Width = 280 
    $DrawSize.Height = 20
    $lblRemovePDs.Size = $DrawSize
    $tbpMoveUsers.Controls.Add($lblRemovePDs)

    #endregion lblRemovePDs

    #region rtbCopyAccess

    $rtbCopyAccess = New-Object System.Windows.Forms.RichTextBox
    $rtbCopyAccess.Name = "rtbCopyAccess"
    $rtbCopyAccess.Text = ""
    $DrawPoint.X = 390
    $DrawPoint.Y = 80
    $rtbCopyAccess.Location = $DrawPoint
    $DrawSize.Width = 280 
    $DrawSize.Height = 150
    $rtbCopyAccess.Size = $DrawSize
    $tbpMoveUsers.Controls.Add($rtbCopyAccess)

    #endregion rtbCopyAccess

    #region lblCopyAccess

    $lblCopyAccess = New-Object System.Windows.Forms.Label
    $lblCopyAccess.Name = "lblCopyAccess"
    $lblCopyAccess.Text = "PDs to copy access from"
    $lblCopyAccess.TextAlign = "BottomCenter" 
    $DrawPoint.X = 390
    $DrawPoint.Y = 60
    $lblCopyAccess.Location = $DrawPoint
    $DrawSize.Width = 280 
    $DrawSize.Height = 20
    $lblCopyAccess.Size = $DrawSize
    $tbpMoveUsers.Controls.Add($lblCopyAccess)

    #endregion lblCopyAccess

    #region rtbMoveResults

    $rtbMoveResults = New-Object System.Windows.Forms.RichTextBox
    $rtbMoveResults.Name = "rtbMoveResults"
    $rtbMoveResults.Text = ""
    $DrawPoint.X = 10
    $DrawPoint.Y = 250
    $rtbMoveResults.Location = $DrawPoint
    $DrawSize.Width = 520
    $DrawSize.Height = 400
    $rtbMoveResults.Size = $DrawSize
    $tbpMoveUsers.Controls.Add($rtbMoveResults)

    #endregion rtbMoveResults

    #region btnUpdateData

    $btnUpdateData = New-Object System.Windows.Forms.Button
    $btnUpdateData.Name = "btnUpdateData"
    $btnUpdateData.Text = "Data date: "+$datadate+"`nUpdate now"
    $DrawPoint.X = 400
    $DrawPoint.Y = 10
    $btnUpdateData.Location = $DrawPoint
    $DrawSize.Width = 130
    $DrawSize.Height = 40
    $btnUpdateData.Size = $DrawSize
    $btnUpdateData.add_Click({btnUpdateDataOnClick})
    $tbpMoveUsers.Controls.Add($btnUpdateData)

    Function btnUpdateDataOnClick {
        $btnUpdateData.Text = "Updating..."
        $form.Refresh()
        $global:hrdata = get-content -Path  $hrpath | Select-Object -skip 1 | ConvertFrom-Csv
        $global:datadate = Get-date -Format "dd/MM/yy"
        $btnUpdateData.Text = "Data date: "+$datadate+"`nUpdate now"
    }
    #endregion btnUpdateData

    

#endregion MovePDUsers

#region RenamePD

    #region btnRenamePD

    $btnRenamePD = New-Object System.Windows.Forms.Button
    $btnRenamePD.Name = "btnRenamePD"
    $btnRenamePD.Text = "Rename"
    $DrawPoint.X = 580
    $DrawPoint.Y = 20
    $btnRenamePD.Location = $DrawPoint
    $DrawSize.Width = 80 
    $DrawSize.Height = 40
    $btnRenamePD.Size = $DrawSize
    $btnRenamePD.add_Click({btnRenamePDOnClick})
    $tbpRenamePD.Controls.Add($btnRenamePD)

    function btnRenamePDOnClick{
        $btnRenamePD.Visible = $False
        RenamePD -newPDtbx $tbxRenameCorrectPD -oldPDtbx $tbxRenameOldPD -userbox $rtbRenameUsers -outbox $rtbRenameResults
       
    }

    #endregion btnRenamePD

    #region btnRenameClearPD

    $btnRenameClearPD = New-Object System.Windows.Forms.Button
    $btnRenameClearPD.Name = "btnRenameClearPD"
    $btnRenameClearPD.Text = "Clear"
    $DrawPoint.X = 580
    $DrawPoint.Y = 80
    $btnRenameClearPD.Location = $DrawPoint
    $DrawSize.Width = 80 
    $DrawSize.Height = 40
    $btnRenameClearPD.Size = $DrawSize
    $btnRenameClearPD.add_Click({btnRenameClearPDOnClick})
    $tbpRenamePD.Controls.Add($btnRenameClearPD)

    function btnRenameClearPDOnClick{
        $tbxRenameCorrectPD.Text = ""
        $tbxRenameOldPD.Text = ""
        $rtbRenameUsers.Text = ""
        $rtbRenameResults.Text = ""
        $btnRenamePD.Visible = $True
    }

    #endregion btnRenameClearPD

    #region lblRenameCorrectPD

    $lblRenameCorrectPD = New-Object System.Windows.Forms.Label
    $lblRenameCorrectPD.Name = "lblRenameCorrectPD"
    $lblRenameCorrectPD.Text = "Correct PD Name"
    $lblRenameCorrectPD.TextAlign = "MiddleCenter" 
    $DrawPoint.X = 20
    $DrawPoint.Y = 20
    $lblRenameCorrectPD.Location = $DrawPoint
    $DrawSize.Width = 520 
    $DrawSize.Height = 20
    $lblRenameCorrectPD.Size = $DrawSize
    $tbpRenamePD.Controls.Add($lblRenameCorrectPD)

    #endregion lblRenameCorrectPD

    #region tbxRenameCorrectPD

    $tbxRenameCorrectPD = New-Object System.Windows.Forms.TextBox
    $tbxRenameCorrectPD.Name = "tbxRenameCorrectPD"
    $tbxRenameCorrectPD.Text = ""
    $DrawPoint.X = 20
    $DrawPoint.Y = 40
    $tbxRenameCorrectPD.Location = $DrawPoint
    $DrawSize.Width = 520 
    $DrawSize.Height = 20
    $tbxRenameCorrectPD.Size = $DrawSize
    $tbpRenamePD.Controls.Add($tbxRenameCorrectPD)

    #endregion tbxRenameCorrectPD
    
    #region lblRenameOldPD

    $lblRenameOldPD = New-Object System.Windows.Forms.Label
    $lblRenameOldPD.Name = "lblRenameOldPD"
    $lblRenameOldPD.Text = "Old PD Name"
    $lblRenameOldPD.TextAlign = "MiddleCenter" 
    $DrawPoint.X = 20
    $DrawPoint.Y = 80
    $lblRenameOldPD.Location = $DrawPoint
    $DrawSize.Width = 520 
    $DrawSize.Height = 20
    $lblRenameOldPD.Size = $DrawSize
    $tbpRenamePD.Controls.Add($lblRenameOldPD)

    #endregion lblRenameOldPD

    #region tbxRenameOldPD

    $tbxRenameOldPD = New-Object System.Windows.Forms.TextBox
    $tbxRenameOldPD.Name = "tbxRenameOldPD"
    $tbxRenameOldPD.Text = ""
    $DrawPoint.X = 20
    $DrawPoint.Y = 100
    $tbxRenameOldPD.Location = $DrawPoint
    $DrawSize.Width = 520 
    $DrawSize.Height = 20
    $tbxRenameOldPD.Size = $DrawSize
    $tbpRenamePD.Controls.Add($tbxRenameOldPD)

    #endregion tbxRenameOldPD 

    #region rtbRenameUsers

    $rtbRenameUsers = New-Object System.Windows.Forms.RichTextBox
    $rtbRenameUsers.Name = "rtbRenameUsers"
    $rtbRenameUsers.Text = ""
    $DrawPoint.X = 10
    $DrawPoint.Y = 140
    $rtbRenameUsers.Location = $DrawPoint
    $DrawSize.Width = 100
    $DrawSize.Height = 200
    $rtbRenameUsers.Size = $DrawSize
    $tbpRenamePD.Controls.Add($rtbRenameUsers)

    #endregion rtbRenameUsers

    #region rtbRenameResults

    $rtbRenameResults = New-Object System.Windows.Forms.RichTextBox
    $rtbRenameResults.Name = "rtbRenameResults"
    $rtbRenameResults.Text = ""
    $DrawPoint.X = 130
    $DrawPoint.Y = 140
    $rtbRenameResults.Location = $DrawPoint
    $DrawSize.Width = 530
    $DrawSize.Height = 200
    $rtbRenameResults.Size = $DrawSize
    $tbpRenamePD.Controls.Add($rtbRenameResults)

    #endregion rtbRenameResults


#endregion RenamePD

$form.AcceptButton = $btnReady["Main1"]
$form.Add_Shown({$form.Activate(); $tbxReady["Main1"].focus()})
$form.ShowDialog()| Out-Null

#region Change Log

#2.0 
#	Loaded tabs and tab controls into hash tables to maximise code re-use 
#	Allowed switching between listbox and textbox views
#	Incorporated compliance checking on access changes.

#2.1
#	had membership view update when changed
#   TODO: changed tab titles to update when objects are loaded.



#endregion change log

