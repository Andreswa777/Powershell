#### connecting the online serviceses
Connect-ExchangeOnline -ShowBanner:$false
Remove-Variable * -ErrorAction SilentlyContinue

Add-Type -AssemblyName System.Windows.Forms

# Create a form
$form = New-Object System.Windows.Forms.Form
$form.Text = "MDO Allow/Block Sender/domain"
$form.Size = New-Object System.Drawing.Size(600, 450)
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
$radioButton1.Text = "Blocked Senders"
$radioButton1.Location = New-Object System.Drawing.Point(310, 60)
$radioButton1.Size = New-Object System.Drawing.Size(500, 20)
$radioButtons += $radioButton1
$radioButton2 = New-Object System.Windows.Forms.RadioButton
$radioButton2.Text = "Blocked Sender Domains"
$radioButton2.Location = New-Object System.Drawing.Point(310, 90)
$radioButton2.Size = New-Object System.Drawing.Size(500, 20)
$radioButtons += $radioButton2
$radioButton3 = New-Object System.Windows.Forms.RadioButton
$radioButton3.Text = "Allowed Senders"
$radioButton3.Location = New-Object System.Drawing.Point(310, 120)
$radioButton3.Size = New-Object System.Drawing.Size(500, 20)
$radioButtons += $radioButton3
$radioButton4 = New-Object System.Windows.Forms.RadioButton
$radioButton4.Text = "Allowed Sender Domain"
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

# Create a help button
$helpButton = New-Object System.Windows.Forms.Button
$helpButton.Location = New-Object System.Drawing.Point(450, 230)
$helpButton.Size = New-Object System.Drawing.Size(100, 20)
$helpButton.Text = "About"
$helpButton.Add_Click({
    # Create a new form for the help popup
    $helpForm = New-Object System.Windows.Forms.Form
    $helpForm.Text = "About"
    $helpForm.Size = New-Object System.Drawing.Size(600, 450)
    $helpForm.StartPosition = "CenterScreen"
    $helpForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
    $helpForm.ControlBox = $false
    
    # Create a label to display the help text
    $helpLabel = New-Object System.Windows.Forms.Label
    $helpLabel.Location = New-Object System.Drawing.Point(10, 10)
    $helpLabel.Size = New-Object System.Drawing.Size(590, 350)

    ####
    $helpLabel.Text = "MDO Allow/Block Sender/domain" + "`n" +
                      "Version: v1.0" + "`n" +
                      "Created by: Andre Swanepoel" + "`n" +
                      "`n" +
                      "Any errors, issues or suggestions, please let me know" + "`n" + "`n" +

                      "Application Function:" + "`n" +
                      "This application allows the user to input multiple lines within the inputText Box" + "`n" +
                      "The user can then add the entries into one of the four actions" + "`n" +
                      "1 - Blocked Sender - Specific email address will be blocked from sending spam to end-users" + "`n" +
                      "2 - Blocked Sender Domains - Complete domain  will be blocked from sending spam to end-users" + "`n" +
                      "3 - Allowed Sender - Specific email address will bypass all the MDO email checks" + "`n" +
                      "4 - Allowed Sender Domains - Complete domain address will bypass all the MDO email checks" + "`n" +
                      "`n" +
                      "This is equivalent of adding items from within the M 365 Defender centre" + "`n" +
                      "Under:" + "`n" +
                      "Email & Collaboration --> Policy & Rules --> Threat Policies --> Anti-Spam --> Anti-spam inbound policy"

    $helpForm.Controls.Add($helpLabel)
    
    # Create a back button to return to the original form
    $backButton = New-Object System.Windows.Forms.Button
    $backButton.Location = New-Object System.Drawing.Point(250, 380)
    $backButton.Size = New-Object System.Drawing.Size(100, 23)
    $backButton.Text = "Back"
    $backButton.Add_Click({
        $helpForm.Close()  # Close the help popup form
    })
    $helpForm.Controls.Add($backButton)
    
    # Show the help popup form as a dialog
    $helpForm.ShowDialog()
})
$form.Controls.Add($helpButton)


# Create a text box for output
$outputTextBox = New-Object System.Windows.Forms.TextBox
$outputTextBox.Multiline = $true
$outputTextBox.Location = New-Object System.Drawing.Point(20, 275)
$outputTextBox.Size = New-Object System.Drawing.Size(550, 130)
$outputTextBox.ScrollBars = "Vertical"
$outputTextBox.ReadOnly = $true
$form.Controls.Add($outputTextBox)

# Submit button click event
$submitButton.Add_Click({
    try {
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
        } else {
            $outputTextBox.Text = "Please select an action and click 'Submit'"
        }
    } catch {
        $outputTextBox.Text = "Error: $($_.Exception.Message)"
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

            try {
                Set-HostedContentFilterPolicy -Identity Default -BlockedSenders $BlockedS -ErrorAction Stop
                
                $output = "The following email addresses were added to BlockedSenders: `r`n"
                $output += "$inputLines `r`n"
                $output += "`r`n"
                $output += "For a full list of currently blocked email addresses, run the following:`r`n"
                $output += "Get-HostedContentFilterPolicy -Identity Default | Select -ExpandProperty BlockedSenders | Out-Gridview"
            } catch {
                $output = "Error: $($_.Exception.Message)"
                $output += "`r`n"
                $output += "`r`n"
                $output += "Accepted input: full email addresses only.`r`n"
                $output += "Example of accepted input:`r`n"
                $output += "Test@test.com"
            }
        }
        "BlockedSenderDomains" {
            $hasInvalidInput = $inputLines -match '@'
            if (-not $hasInvalidInput) {
                $BlockedD = @{Add = $inputLines}
                Try {
                    Set-HostedContentFilterPolicy -Identity Default -BlockedSenderDomains $BlockedD -ErrorAction Stop
                    $output = "The following domains were added to BlockedSenderDomains:`r`n"
                    $output += "$inputLines `r`n"
                    $output += "`r`n"
                    $output += "For a full list of currently blocked domains, run the following:`r`n"
                    $output += "Get-HostedContentFilterPolicy -Identity Default | Select -ExpandProperty BlockedSenderDomains | Out-Gridview"
                } Catch {
                    $output = "Error: $($_.Exception.Message)`r`n"
                    $output += "`r`n"
                    $output += "Accepted input: Domain name without the @.`r`n"
                    $output += "Example of accepted input:`r`n"
                    $output += "test.com"
                }
            } else {
                $output = "Error: At least one of the input lines contains the '@' symbol. Please provide valid domain names.`r`n"
                $output += "`r`n"
                $output += "Accepted input: Domain name without the @.`r`n"
                $output += "Example of accepted input:`r`n"
                $output += "test.com"
            }
        }
        "AllowedSenders" {
            $AllowedS = @{Add=$inputLines}            

            try {
                Set-HostedContentFilterPolicy -Identity Default -AllowedSenders $AllowedS -ErrorAction Stop
                
                $output = "The following email addresses were added to AllowedSenders: `r`n"
                $output += "$inputLines `r`n"
                $output += "`r`n"
                $output += "For a full list of currently allowed email addresses, run the following:`r`n"
                $output += "Get-HostedContentFilterPolicy -Identity Default | Select -ExpandProperty AllowedSenders | Out-Gridview"
            } catch {
                $output = "Error: $($_.Exception.Message)"
                $output += "`r`n"
                $output += "`r`n"
                $output += "Accepted input: full email addresses only.`r`n"
                $output += "Example of accepted input:`r`n"
                $output += "Test@test.com"
            }
        }
        "AllowedSenderDomains" {
            $hasInvalidInput = $inputLines -match '@'
            if (-not $hasInvalidInput) {
                $AllowedD = @{Add = $inputLines}
                try {
                    Set-HostedContentFilterPolicy -Identity Default -AllowedSenderDomains $AllowedD -ErrorAction Stop
                    $output = "The following domains were added to AllowedSenderDomains:`r`n"
                    $output += "$inputLines `r`n"
                    $output += "`r`n"
                    $output += "For a full list of currently allowed domains, run the following:`r`n"
                    $output += "Get-HostedContentFilterPolicy -Identity Default | Select -ExpandProperty AllowedSenderDomains | Out-Gridview"
                } catch {
                    $output = "Error: $($_.Exception.Message)`r`n"
                    $output += "`r`n"
                    $output += "Accepted input: Domain name without the @.`r`n"
                    $output += "Example of accepted input:`r`n"
                    $output += "test.com"
                }
            } else {
                $output = "Error: At least one of the input lines contains the '@' symbol. Please provide valid domain names.`r`n"
                $output += "`r`n"
                $output += "Accepted input: Domain name without the @.`r`n"
                $output += "Example of accepted input:`r`n"
                $output += "test.com"
            }
        }
    }
    
    # Update the output in the text box
    $outputTextBox.Text = $output
}

# Show the form
$form.ShowDialog()