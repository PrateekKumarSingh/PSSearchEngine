Function Invoke-BingAutoComplete
{
    Return (Invoke-RestMethod -Uri "http://api.bing.com/qsml.aspx?query=$($TextBox1.text)").searchsuggestion.section.item.text
}

#Function to fetch the data from Wolfram|Alpha API based on user query
Function Invoke-WolframAlphaAPI($Query)
{
Return (Invoke-RestMethod -Uri "http://api.wolframalpha.com/v2/query?appid=46XTUT-6T5H7K4V32&input=$($Query.Replace(' ','%20'))").queryresult
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
                              }

$DidYouMeanEventHandler =[System.EventHandler]{
                                $TextBox1.Text = $DidYouMeanText
                                $DidYouMeanButton.Enabled = $False                                
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
    
    #Define the Progress Bar
    $ProgressBar = New-Object System.Windows.Forms.ProgressBar
    $ProgressBar.Maximum = 100
    $ProgressBar.Minimum = 0
    $ProgressBar.Height = 10
    $ProgressBar.Width = 430
    $ProgressBar.ForeColor = 'Blue'
    $ProgressBar.Style = 'block'
    
    #Define the Form
    $Form = New-Object system.Windows.Forms.Form
    $Form.Text="Search your Query here  [Powered by Wolfram|Alpha API]"
    $Form.BackColor = 'white'
    $Form.AutoSize = $False
    $Form.MinimizeBox = $False
    $Form.MaximizeBox = $False
    $Form.WindowState = "Normal"
    $Form.StartPosition = "CenterScreen"
    $Form.Height = 500
    $Form.Width = 470
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
Function Create-PanelStructure($Result)
{
    #Try
    #{
        If($Result.success -eq $True)
        {

        #Formula to calculate Progress bar increment each time a Sub Pod is parsed
        $Increment = (100/[int]$Result.numpods)

        $i=0 #Initialize ProgressBar Value 

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
            Write-host $i

            }
        }
        Else
        {
            $DidYouMeans =  $Result.didyoumeans.didyoumean
            
            Foreach($DidYouMean in $DidYouMeans)
            {
                $DidYouMeanButton = New-Object System.Windows.Forms.Button
                $DidYouMeanText = $DidYouMean."#text"
                $DidYouMeanButton.Text = "Did you mean $DidYouMeanText"
                $DidYouMeanButton.AutoSize = $True
                #$Label.Font = $Font
                #$Label.ForeColor = 'Blue'
                $DidYouMeanButton.Add_Click($DidYouMeanEventHandler)
                $Panel2.Controls.Add($DidYouMeanButton)
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