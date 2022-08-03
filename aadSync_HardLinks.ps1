 <#

     DISCLAIMER:
        Run this at your own risk. If you do not have a solid understanding of AzureAD Sync AND ImmutableID's please read through the Microsoft Documentation first.
        This is a modified version of the script posted by "D.G." here: https://community.spiceworks.com/how_to/122371-hard-link-ad-accounts-to-existing-office-365-users-when-soft-match-fails

     ORIGINAL SCRIPT:
        do{
        $ADGuidUser = Get-ADUser -Filter * | Select Name,ObjectGUID | Sort-Object Name | Out-GridView -Title "Select Local AD User To Get Immutable ID for" -PassThru
        $UserimmutableID = [System.Convert]::ToBase64String($ADGuidUser.ObjectGUID.tobytearray())
        $OnlineUser = Get-MsolUser | Select UserPrincipalName,DisplayName,ProxyAddresses,ImmutableID | Sort-Object DisplayName | Out-GridView -Title "Select The Office 365 Online User To HardLink The AD User To" -PassThru
        Set-MSOLuser -UserPrincipalName $OnlineUser.UserPrincipalName -ImmutableID $UserimmutableID
        $Repeat = read-host Do you want to choose another user? Y or N } while ($Repeat -eq "Y")

     AUTHORS NOTE
        Last updated 08/03/2022 by Erick Powers
        Feel free to take this and do anything you want with it. 
        Even with the modifications I thank D.G. for this.

 #>

# Popup Prompt Title 
$PopupTitle = "AzureAD Sync - Hard Link"

# Popup Text that will ask if you want to repeat.
$PopupBodyRepeat = "Do you want to Hard Link another user?"

# Popup Text for the Alert Box that lets you know that you chose to not apply the ImmutableID.
$PopupBodyNoActionTaken = "You have chosen to not apply your selections."

# Popup options presented to the user, default is no... so if someone just hits enter its a No... you can press 'Y' or 'N' and the selection will be made.
$options = [System.Management.Automation.Host.ChoiceDescription[]] @("&Yes", "&No")

# Set the default on all prompts to No ... in case of accidents.
[int]$PopupDefaults = 1

# Lets... Do this? 
do
{
    # This pulls the list of AD Users for you to select from.
    $ADGuidUser = Get-ADUser -Filter * | 
        Select Name,ObjectGUID | 
            Sort-Object Name | 
                Out-GridView -Title "Select Local AD User To Get Immutable ID for then press Enter." -PassThru
    
    # This is used to display the selected user in the MsolUser Selection Prompt.
    $SelectedUserName = $ADGuidUser.Name
    
    # This converts the ObjectGUID of your AD User to the ImmutableID that will be in place on AzureAD
    $UserimmutableID = [System.Convert]::ToBase64String($ADGuidUser.ObjectGUID.tobytearray())
    
    # This pulls the list of AzureAD Users for you select and apply the ImmutableID to. 
    $OnlineUser = Get-MsolUser | 
        Select UserPrincipalName,DisplayName,ProxyAddresses,ImmutableID | 
            Sort-Object DisplayName | 
                Out-GridView -Title "Select The Office 365 Online User To HardLink The AD User To, then press Enter. You selected the following user from your AD: $SelectedUserName " -PassThru
    
    # This if for displaying the Online User you choose.
    $SelectedOnlineUser = $OnlineUser.UserPrincipalName
    
    # Defined this one here because... 
    $PopupBodyConfirmSelections = "You have selected $SelectedUserName from your local Active Directory Server. `nYou have selected $SelectedOnlineUser from AzureAD. `nDo you wish to continue? (Y/N)"
    
    # pop the prompt for confirmation on the two items selected. if we say yes to process it sets the ImmutableID on the AzureAD Profile.
    $ConfirmationPrompt = $host.UI.PromptForChoice($PopupTitle , $PopupBodyConfirmSelections , $options, $PopupDefaults)
    switch($ConfirmationPrompt)
    {
        0 { 
            Set-MSOLuser -UserPrincipalName $OnlineUser.UserPrincipalName -ImmutableID $UserimmutableID 
          }
        1 { 
            $Repeat = "N"
            [System.Windows.Forms.MessageBox]::Show($PopupBodyNoActionTaken,'WARNING')
          }
    }

    # Pop the run another prompt unless we said no to the selection confirmations.
    if($Repeat -ne "N")
    {
        $RepeatPrompt = $host.UI.PromptForChoice($PopupTitle , $PopupBodyRepeat , $options, $PopupDefaults)
        switch($RepeatPrompt)
        {
            0 { $Repeat = "Y" }
            1 { $Repeat = "N" }
        }
    }
} 
while ($Repeat -eq "Y")
