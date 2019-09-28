Import-Module SkypeOnlineConnector
$Cred = Get-Credential
$CSSession = New-CsOnlineSession -Credential $Cred
Import-PSSession -Session $CSSession




# Specifcy info about new employees
# normally this would be done by polling the inboind info from AD

$Region = "NOAM"
$Country = "US"
$Area = Read-Host -Prompt "Enter the state Abreviation here, example Colorado is CO" 
$CityCode = Read-Host -Prompt "Enter the city code here, This varies by state. Ex.Denver, CO is NOAM-US-CO-DE" 
$NewPeople = Read-Host -Prompt "Enter the list of voice users here separated by a comma" 
$officeLoc = Read-Host -Prompt "Enter the City Name of your Emergency Location here. If the city is two words, just put the first word. Ex. the City is Greenwood Village, you will put Greenwood" 

#Do we have number in location of new employees? if not, get some

$NumAvailable = (Get-CsOnlineTelephoneNumber -capitalormajorcity $citycode).count
$Numneeded = $NewPeople.count - $NumAvailable

If ($Numneeded > 0)
{
    #
    # Get Numbers in the location we need.
    #Need to search inventory first, then select the numbers from the reserved inventory}
    #
    $SearchOut = search-csonlinetelephonenumberinventory -inventorytype subscriber -regionalgroup $Region -CountryorRegion $Country -Area $Area
    Select-Csonlinetelephonenumberinventory -reservationID $searchout.ReservationID -TelephoneNumbers $searchout.reservations.numbers.number -Region $Region

}
#Get the emergency services location for these people

$locID = (Get-CSonlineLisCivicAddress -City $officeLoc).DefaultlocationId

# Get available numbers in that region

$NewNums = Get-CSonlinetelephonenumber -capitalormajorcity $citycode -resultsize $NewPeople.Count

#Apply to Each user
#
for($i=0;$i -lt $Newpeople.count;$i++){
    Set-CsonlineVoiceUser -Identity $Newpeople -locationID $locID -TelephoneNumber $NewNums;
    Get-CSonlineVoiceUser -Identity $NewPeople
}