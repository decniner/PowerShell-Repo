#Contact decniner@gmail.com
#Here is an example of a PowerShell script that accepts input in JSON format:

# Parse the input JSON string
$input = Get-Content -Raw | ConvertFrom-Json

# Extract the values from the input object
$value1 = $input.Value1
$value2 = $input.Value2

# Do something with the values
Write-Output "Value 1 is $value1 and value 2 is $value2"

<#
This script reads the input from standard input (using Get-Content -Raw) and converts it to a PowerShell object using ConvertFrom-Json. It then extracts the values Value1 and Value2 from the input object and stores them in variables. Finally, it writes the values to standard output.
To pass input to this script, you can use the | (pipe) operator to pipe the output of a command to the script. For example:
#>

echo '{"Value1": "abc", "Value2": 123}' | .\myscript.ps1

<#
This will pass the JSON string {"Value1": "abc", "Value2": 123} to the script as input, which will parse it and output Value 1 is abc and value 2 is 123.
#>
