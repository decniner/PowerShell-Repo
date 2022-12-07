# example.ps1

# Import the necessary modules
Import-Module WebRequest

# Set the API endpoint URL
$apiUrl = "http://moonshot-server/api/v1/blades"

# Set the API authentication credentials
$cred = Get-Credential

# Set the options for the API request
$options = @{
    Method = "GET"
    Credential = $cred
}

# Send the API request and save the response
$response = Request-WebRequest $apiUrl $options

# Convert the response to a PowerShell object
$bladeStatus = ConvertFrom-Json $response.Content

# Create an HTML table to display the blade status information
$table = "<table>
              <tr>
                  <th>Blade ID</th>
                  <th>Status</th>
                  <th>Health</th>
              </tr>"

# Loop through each blade in the response
foreach ($blade in $bladeStatus)
{
    # Create a row in the HTML table for the current blade
    $table += "<tr>
                   <td>$($blade.id)</td>
                   <td>$($blade.status)</td>
                   <td>$($blade.health)</td>
               </tr>"
}

# Close the HTML table
$table += "</table>"

# Save the HTML table to a file
$table | Out-File "blade-status.html"
