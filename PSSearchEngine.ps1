$Global:SearchedFlag = $False

Function Invoke-BingAutoComplete
{
    Return (Invoke-RestMethod -Uri "http://api.bing.com/qsml.aspx?query=$($TextBox1.text)").searchsuggestion.section.item.text
}

#Function to fetch the data from Wolfram|Alpha API based on user query
Function Invoke-WolframAlphaAPI($Global:Query)
{
Return (Invoke-RestMethod -Uri "http://api.wolframalpha.com/v2/query?appid=46XTUT-6T5H7K4V32&input=$($Query.Replace(' ','%20'))").queryresult
}

#Extract Results to HTML file
Function Get-Html($R)
{

    "<html>"
    "<body>"            
    Foreach($p in $Result.pod)
    {
        $subpod = $p.subpod
        
    
        "<h3>$($p.title)</h3>"
    
            foreach($s in $subpod)
            {
                #Incase plain text field is blank, display the image in the panel
                if($s.plaintext -eq '')
                {
                    
                     "<img src='$($s.img.src)' />"
                }
                Else
                {
                     "<p>$($s.plaintext)</p>"
                }
            }
           #"<hr>"
    }
    "</body>"
    "</html>"
}

#Eventhandler and Flow control once the Search button is pressed
$EventHandler =[System.EventHandler]{

                                $Panel2.Visible = $False
                                $Panel2.Controls.clear()
                                $ProgressBar.value = 0
                                $Panel3.Visible = $True
                                $Button.Enabled = $False                                
                                Create-PanelStructure $(Invoke-WolframAlphaAPI $TextBox1.Text)
                                $Panel2.Visible = $True
                                $Button.Enabled = $True
                                $Panel3.Visible = $False
                                $Global:SearchedFlag = $True
                              }

$SaveEventHandler = [System.EventHandler]{

    Get-Html $result | Out-File "$env:TEMP\$Query.html"
    ii "$env:TEMP\$Query.html"
}

$DidYouMeanEventHandler =[System.EventHandler]{
                                $Panel2.Visible = $False
                                $Panel2.Controls.clear()
                                $TextBox1.Text = $DidYouMeanText
                                $DidYouMeanButton.visible = $False                                
                                $ProgressBar.value = 0
                                $Panel3.Visible = $True
                                Create-PanelStructure $(Invoke-WolframAlphaAPI $DidYouMeanText)
                                $Panel2.Visible = $True
                                $Button.Enabled = $True
                                $Panel3.Visible = $False
                              }

#Funtion to Create the Basic form and its Structure.
Function Create-WindowsForm()
{    
    #Calling the Assemblies
    [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
    #[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")

    #Define Text Font object
    $Font = New-Object System.Drawing.Font("lucida sans",10,[System.Drawing.FontStyle]::bold) 
    $Font2 = New-Object System.Drawing.Font("lucida sans",13,[System.Drawing.FontStyle]::bold) 
                
    #Define TextBox1 for input
    $TextBox1 = New-Object “System.Windows.Forms.RichTextBox”;
    $TextBox1.BorderStyle = 'fixed3d'
    $TextBox1.BackColor = 'snow'
    $TextBox1.Left = 10;
    $TextBox1.Top = 10;
    $TextBox1.Height = 40
    $TextBox1.width = 340;
    $TextBox1.Font = $Font2
    
    $TextBox1.add_keyup({     
                            If($Global:SearchedFlag -eq $true)
                            {
                                $SearchedFlag
                                $Panel2.controls.clear()
                                $Global:SearchedFlag = $False     
                            }
                            $Data = Invoke-BingAutoComplete
                            $Data | %{$StrWithLineBreaks+=$_+';'}
                            $AutocompleteLabel.text=$StrWithLineBreaks -replace ";","`n"
    }) 

    #Define Search Button
    $Button = New-Object System.Windows.Forms.Button
    $Button.Text = "Search"
    $Button.Font = $Font2
    $Button.Height = 40
    $Button.Add_Click($EventHandler)

    #Define Save Button
    $SaveButton = New-Object System.Windows.Forms.Button
    $SaveButton.text = "Save"    
    $SaveButton.Font = $Font2
    $SaveButton.Height =  40
    $SaveButton.Add_Click($SaveEventHandler)

    #Define the Progress Bar
    $ProgressBar = New-Object System.Windows.Forms.ProgressBar
    $ProgressBar.Maximum = 100
    $ProgressBar.Minimum = 0
    $ProgressBar.Height = 10
    $ProgressBar.Width = 500
    $ProgressBar.ForeColor = 'Blue'
    $ProgressBar.Style = 'block'
    $ProgressBar.Visible = $true
    
    #Define the Form
    $Form = New-Object system.Windows.Forms.Form
    $Form.Text="PS Search Engine"
    $Form.BackColor = 'white'
    $Form.AutoSize = $False
    $Form.MinimizeBox = $False
    $Form.MaximizeBox = $False
    $Form.WindowState = "Normal"
    $Form.StartPosition = "CenterScreen"
    $Form.Height = 500
    $Form.Width = 550
    $Form.Icon = [System.Drawing.Icon]::ExtractAssociatedIcon((Get-Process -id $pid | Select-Object -ExpandProperty Path))                    
    $Form.AutoScroll = $True
    $Form.AcceptButton = $Button

    $AutocompleteLabel = New-Object System.Windows.Forms.Label
    $AutocompleteLabel.AutoSize = $True
    $AutocompleteLabel.Font = $Font
    
    #Define the Base Panel on which we'll add 3 sub panels
    $RootPanel = new-object System.Windows.Forms.FlowLayoutPanel
    $RootPanel.AutoSize = $True
    $RootPanel.FlowDirection = 'topdown'
    
    #Define Sub Panel 1
    $Panel1 = new-object System.Windows.Forms.FlowLayoutPanel
    $Panel1.AutoSize = $True
    $Panel1.Controls.Add($TextBox1)
    $Panel1.Controls.Add($Button)
    $Panel1.Controls.Add($SaveButton)


    
    #Define Sub Panel 2
    $Panel2 = new-object System.Windows.Forms.FlowLayoutPanel
    $Panel2.AutoSize = $True
    $Panel2.FlowDirection = 'topdown'
        
    #To adjust output Panel size accordint to maximum sizes, to avoid data or image getting cropped.
    $Panel2.Width = ($Panel2.Controls.width | measure -Maximum).maximum
    $Panel2.Height = ($Panel2.Controls.Height | Measure -Sum).sum + 50

    $Panel2.Controls.Add($AutocompleteLabel)

    #Define Sub Panel 3
    $Panel3 = new-object System.Windows.Forms.FlowLayoutPanel
    $Panel3.AutoSize = $True
    $Panel3.Visible = $False
    $Panel3.Controls.Add($ProgressBar)

    #Add all panels to the root Panel, so that the flow direction is Top to Down.
    $RootPanel.Controls.Add($Panel1)
    $RootPanel.Controls.Add($Panel3)
    $RootPanel.Controls.Add($Panel2)
    
    #Add Root Panel to the Form and display it.
    $Form.Controls.Add($RootPanel)
    [void]$Form.ShowDialog()

    
        
}

#Function to Create the data structure for Output on Panel 3
Function Create-PanelStructure($Global:Result)
{
    #Try
    #{
        If($Result.success -eq $True)
        {

        #Formula to calculate Progress bar increment each time a Sub Pod is parsed
        $Increment = (100/[int]$Result.numpods)

        $i=0 #Initialize ProgressBar Value 


        If($Result.warnings.spellcheck.text)
        {
            $WarningLabel = New-Object Windows.forms.label
            $WarningLabel.Text = "Warning : $($Result.warnings.spellcheck.text)"
            $WarningLabel.Font = $Font
            $WarningLabel.ForeColor = "Red"
            $WarningLabel.AutoSize = $True
            $Panel2.Controls.Add($WarningLabel)
        }
         

            Foreach($p in $Result.pod)
            {
                $subpod = $p.subpod
                
                #Create new Label for all POD Titles
                $LabelTitle = New-Object System.Windows.Forms.Label
                $LabelTitle.AutoSize = $True
                $LabelTitle.Text = ($P.title).toUpper()
                $LabelTitle.Font = $Font
                $Panel2.Controls.Add($LabelTitle)
                
                #If($P.title -like "*Defini*")
                #{
                #    
                #}

                #



                    foreach($s in $subpod)
                    {
                        #Incase plain text field is blank, display the image in the panel
                        if($s.plaintext -eq '')
                        {
                             #Create new PictureBox for all Sub POD Images
                             $pictureBox = new-object Windows.Forms.PictureBox
                             $pictureBox.Load($s.img.src)
                             $pictureBox.SizeMode = 'AutoSize'
                             $Panel2.controls.add($pictureBox)                    
                        }
                        Else
                        {
                             #Create new Label for all Sub POD plain text
                             $Label = New-Object Windows.forms.label
                             $Label.AutoSize = $True
                             $Label.Text = $s.plaintext
                             $Panel2.Controls.Add($Label)
                        }
                    }

            #Increment the ProgressBar and display increasing values
            $i=$i+$Increment
            $ProgressBar.Value = $i
            #Write-host $i

            }
        }
        ElseIf($Result.didyoumeans.didyoumean)
        {
            $DidYouMeans =  $Result.didyoumeans.didyoumean
            $didyoumeans
            
            Foreach($DidYouMean in $DidYouMeans)
            {
                $GLobal:DidYouMeanButton = New-Object System.Windows.Forms.Button
                $Global:DidYouMeanText = $DidYouMean."#text"
                $DidYouMeanButton.Text = "$DidYouMeanText"
                $DidYouMeanButton.AutoSize = $True
                $DidYouMeanButton.ForeColor = "White"
                $DidYouMeanButton.BackColor = "Black"
                $DidYouMeanButton.Font = $Font
                $DidYouMeanButton
                
                $Global:DidYouMeanLabel = New-Object System.Windows.Forms.Label
                $DidYouMeanLabel.Font = $Font
                $DidYouMeanLabel.Text = "Did you mean ?"
                $DidYouMeanLabel.AutoSize = $True

                $DidYouMeanButton.Add_Click($DidYouMeanEventHandler)
                $Panel2.Controls.Add($DidYouMeanLabel)
                $Panel2.Controls.Add($DidYouMeanButton)
            }
        }
        ElseIf($Result.tips.tip)
        {
                $Tips =  $Result.Tips.Tip
            
                Foreach($Tip in $Tips)
                {
                    
                    $Global:TipsLabel = New-Object System.Windows.Forms.Label
                    $TipsLabel.Font = $Font
                    $TipsLabel.Text = $tip.Text
                    $TipsLabel.AutoSize = $True
                    $Panel2.Controls.Add($tipsLabel)
                }
        }
    #}
    #catch
    #{
    #    $Label = New-Object System.Windows.Forms.Label
    #    $Label.Text = "Something went wrong, Please close the window and try again"
    #    $Label.AutoSize = $True
    #    $Label.Font = $Font
    #    $Label.ForeColor = 'red'
    #    $Panel2.Controls.Add($Label)
    #
    #}
}

#Calling the Function to start the tool
Create-WindowsForm
