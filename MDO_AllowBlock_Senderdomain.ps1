#### connecting the online serviceses
Connect-ExchangeOnline -ShowBanner:$false
Remove-Variable * -ErrorAction SilentlyContinue

Add-Type -AssemblyName System.Windows.Forms

# Create a form
$form = New-Object System.Windows.Forms.Form
$form.Text = "MDO Allow/Block Sender/domain"
$form.Size = New-Object System.Drawing.Size(800, 450)
$form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::Fixed3D
$form.MaximizeBox = $false

# Create label for inputTextBox
$labelInputRadio = New-Object System.Windows.Forms.Label
$labelInputRadio.Text = "Select Action:"
$labelInputRadio.Location = New-Object System.Drawing.Point(310, 30)
$labelInputRadio.Size = New-Object System.Drawing.Size(200, 20)
$labelInputRadio.Font = New-Object System.Drawing.Font($labelInputRadio.Font.Name, $labelInputRadio.Font.Size, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($labelInputRadio)

# Create a radio button group
$radioButtons = @()
$radioButton1 = New-Object System.Windows.Forms.RadioButton
$radioButton1.Text = "Blocked Senders - (full email address) example of correct input eg. test@test.com"
$radioButton1.Location = New-Object System.Drawing.Point(310, 60)
$radioButton1.Size = New-Object System.Drawing.Size(500, 20)
$radioButtons += $radioButton1
$radioButton2 = New-Object System.Windows.Forms.RadioButton
$radioButton2.Text = "Blocked Sender Domains - (Do not include the @) example of correct input: test.com"
$radioButton2.Location = New-Object System.Drawing.Point(310, 90)
$radioButton2.Size = New-Object System.Drawing.Size(500, 20)
$radioButtons += $radioButton2
$radioButton3 = New-Object System.Windows.Forms.RadioButton
$radioButton3.Text = "Allowed Senders (full email address) example of correct input eg. test@test.com"
$radioButton3.Location = New-Object System.Drawing.Point(310, 120)
$radioButton3.Size = New-Object System.Drawing.Size(500, 20)
$radioButtons += $radioButton3
$radioButton4 = New-Object System.Windows.Forms.RadioButton
$radioButton4.Text = "Allowed Sender Domain - (Do not include the @) example of correct input: test.com"
$radioButton4.Location = New-Object System.Drawing.Point(310, 150)
$radioButton4.Size = New-Object System.Drawing.Size(500, 20)
$radioButtons += $radioButton4
$form.Controls.AddRange($radioButtons)

# Create label for inputTextBox
$labelInput = New-Object System.Windows.Forms.Label
$labelInput.Text = "Enter domain or sender address: (Separate lines)"
$labelInput.Location = New-Object System.Drawing.Point(20, 10)
$labelInput.Size = New-Object System.Drawing.Size(300, 20)
$labelInput.Font = New-Object System.Drawing.Font($labelInput.Font.Name, $labelInput.Font.Size, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($labelInput)

# Create a multiline input box
$inputTextBox = New-Object System.Windows.Forms.TextBox
$inputTextBox.Multiline = $true
$inputTextBox.Location = New-Object System.Drawing.Point(20, 30)
$inputTextBox.Size = New-Object System.Drawing.Size(270, 180)
$inputTextBox.ScrollBars = "Vertical"
$inputTextBox.Anchor = "Left, Top"
$form.Controls.Add($inputTextBox)


# Create a submit button
$submitButton = New-Object System.Windows.Forms.Button
$submitButton.Text = "Submit"
$submitButton.Location = New-Object System.Drawing.Point(150, 230)
$submitButton.Size = New-Object System.Drawing.Size(100, 20)
$form.Controls.Add($submitButton)

# Create an exit button
$exitButton = New-Object System.Windows.Forms.Button
$exitButton.Text = "Exit"
$exitButton.Location = New-Object System.Drawing.Point(250, 230)
$exitButton.Size = New-Object System.Drawing.Size(100, 20)
$form.Controls.Add($exitButton)

# Create a text box for output
$outputTextBox = New-Object System.Windows.Forms.TextBox
$outputTextBox.Multiline = $true
$outputTextBox.Location = New-Object System.Drawing.Point(20, 275)
$outputTextBox.Size = New-Object System.Drawing.Size(750, 130)
$outputTextBox.ScrollBars = "Vertical"
$outputTextBox.ReadOnly = $true
$form.Controls.Add($outputTextBox)

# Submit button click event
$submitButton.Add_Click({
    # Check if connected to Exchange
    if ($radioButton1.Checked) {
            # BlockedSenders selected
            addExclusion "BlockedSenders"
        } elseif ($radioButton2.Checked) {
            # BlockedSenderDomains selected
            addExclusion "BlockedSenderDomains"
        } elseif ($radioButton3.Checked) {
            # AllowedSenders selected
            addExclusion "AllowedSenders"
        } elseif ($radioButton4.Checked) {
            # AllowedSenderDomains selected
            addExclusion "AllowedSenderDomains"
        }else
        {
            $outputTextBox.Text = "Please select an action and click 'Submit'"
        }
})

# Exit button click event
$exitButton.Add_Click({
    # Disconnect from Exchange
    Disconnect-ExchangeOnline -Confirm:$false

    # Close the GUI
    $form.Close()
})

# Function to add exclusion
function addExclusion($exclusionType) {
    $policyIdentity = "Default"
    $outputTextBox.Clear()

    $output = ""

    $inputText = $inputTextBox.Text
    $inputLines = $inputText -split "`r?`n" | Where-Object { $_ -ne '' }
    
    switch ($exclusionType) {
        "BlockedSenders" {
            $BlockedS = @{Add=$inputLines}            

            Set-HostedContentFilterPolicy -Identity Default -BlockedSenders $BlockedS
            
            $output = "The following email addresses were added to BlockedSenders: `r`n"
            $output += "$inputLines `r`n"
            $output += "`r`n"
            $output += "For a full list of currently blocked email addresses, run the following:`r`n"            $output += "Get-HostedContentFilterPolicy -Identity Default | Select -ExpandProperty BlockedSenders | Out-Gridview"
        }
        "BlockedSenderDomains" {
            # Add your command for BlockedSenderDomains here
            $BlockedD = @{Add=$inputLines}            

            Set-HostedContentFilterPolicy -Identity Default -BlockedSenderDomains $BlockedD
            
            $output = "The following domains were added to BlockedSenderDomains: `r`n"
            $output += "$inputLines `r`n"
            $output += "`r`n"
            $output += "For a full list of currently blocked domains, run the following:`r`n"            $output += "Get-HostedContentFilterPolicy -Identity Default | Select -ExpandProperty BlockedSenderDomains | Out-Gridview"
            }

        "AllowedSenders" {
            # Add your command for AllowedSenders here
            $AllowedS = @{Add=$inputLines}            

            Set-HostedContentFilterPolicy -Identity Default -AllowedSenders $AllowedS
            
            $output = "The following email addresses were added to AllowedSenders: `r`n"
            $output += "$inputLines `r`n"
            $output += "`r`n"
            $output += "For a full list of currently allowed email addresses, run the following:`r`n"            $output += "Get-HostedContentFilterPolicy -Identity Default | Select -ExpandProperty AllowedSenders | Out-Gridview"
        }
        "AllowedSenderDomains" {
            # Add your command for AllowedSenderDomains here
            $AllowedD = @{Add=$inputLines}            

            Set-HostedContentFilterPolicy -Identity Default -AllowedSenderDomains $AllowedD
            
            $output = "The following domains were added to AllowedSenderDomains: `r`n"
            $output += "$inputLines `r`n"
            $output += "`r`n"
            $output += "For a full list of currently allowed domains, run the following:`r`n"            $output += "Get-HostedContentFilterPolicy -Identity Default | Select -ExpandProperty AllowedSenders | Out-Gridview"
        }
    }
    
    # Update the output in the text box
    $outputTextBox.Text = $output
}

# Show the form
$form.ShowDialog()