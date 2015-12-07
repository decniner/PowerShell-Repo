function Get-ComputerName{
	if($txtComputer.Text -eq "." -or $txtComputer.Text -eq "localhost" -or $txtComputer.Text -eq "" -or $txtComputer.Text -eq $null){$txtComputer.Text = hostname}
	$Script:ComputerName = $txtComputer.Text.Trim()
}
