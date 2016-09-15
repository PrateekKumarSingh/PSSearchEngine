#region Event Handlers
    
    #Event handlers
    $EventHandler =[System.EventHandler]{
    
                                    $Panel2.Visible = $False
                                    $Panel2.Controls.clear()
                                    $StatusPanel.Controls.clear()
                                    $StatusPanel.controls.add($ProgressBar)
                                    $ProgressBar.value = 0
                                    $StatusPanel.Visible = $True
                                    $Button.Enabled = $False                                
                                    DisplayResults $(Invoke-WolframAlphaAPI $TextBox1.Text)
                                    $Panel2.Visible = $True
                                    $Button.Enabled = $True
                                    $StatusPanel.controls.remove($ProgressBar)
                                  }
    
    $SaveEventHandler = [System.EventHandler]{
    
                                    Get-Html $result | Out-File "$env:TEMP\$Query.html"
                                    $StatusPanel.controls.clear()
                                    $StatusPanel.controls.add($StatusLabel)
                                    $StatusLabel.text = "Saved as File : "
                                    $LinkLabel = New-Object System.Windows.Forms.LinkLabel
                                    $LinkLabel.Text = "$env:TEMP\$Query.html  "
                                    $LinkLabel.AutoSize = $true
                                    $LinkLabel.Font = $ItalicFont
                                    $LinkLabel.add_Click({Invoke-Item "$env:TEMP\$Query.html"})
                                    $StatusPanel.Controls.Add($LinkLabel)
                                    #$OpenButton =  new-object System.Windows.Forms.Button
                                    #$OpenButton.Text = "Open"
                                    #$OpenButton.AutoSize = $True
                                    #$OpenButton.ForeColor = "White"
                                    #$OpenButton.BackColor = "Black"
                                    #$OpenButton.Font = $Font
                                    #$OpenButton.Add_Click({Invoke-Item "$env:TEMP\$Query.html"})
                                    #$StatusPanel.controls.add($OpenButton)
                                    #ii "$env:TEMP\$Query.html"
    }
    
    $DidYouMeanEventHandler =[System.EventHandler]{
                                    
                                    $Panel2.Visible = $False
                                    $Panel2.Controls.clear()
                                    $TextBox1.Text = $DidYouMeanText
                                    $DidYouMeanButton.visible = $False 
                                    $StatusPanel.Controls.clear()                               
                                    $ProgressBar.value = 0
                                    $StatusPanel.Visible = $True
                                    DisplayResults $(Invoke-WolframAlphaAPI $DidYouMeanText)
                                    $Panel2.Visible = $True
                                    $Button.Enabled = $True
                                    $StatusPanel.controls.remove($ProgressBar)
                                  }.GetNewClosure()
    
    $AutoCompleteKeyupEventhandler =  [System.Windows.Forms.KeyEventHandler]{                           
                                    
                                    $Panel2.Controls.clear()
                                    $StatusPanel.Controls.clear()
                                    $StrWithLineBreaks=@()
                                    $Data = Invoke-BingAutoComplete
                                    $Data | %{$StrWithLineBreaks+=$_+';'}
                                    $AutocompleteLabel.text=$StrWithLineBreaks -replace ";","`n"
                                    $Panel2.Controls.Add($AutocompleteLabel)
    }

#endregion Event Handlers

#region Variable Definition

    #Calling the Assemblies
    [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
    #[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")
    
    #Define Text Font object
    $ItalicFont = New-Object System.Drawing.Font("lucida sans",8,[System.Drawing.FontStyle]::Italic) 
    $Font = New-Object System.Drawing.Font("lucida sans",10,[System.Drawing.FontStyle]::bold) 
    $Font2 = New-Object System.Drawing.Font("lucida sans",13,[System.Drawing.FontStyle]::bold) 
    
    #Define TextBox1 for input
    $TextBox1 = New-Object “System.Windows.Forms.RichTextBox”;
    
    $TextBox1.BorderStyle = 'fixed3d'
    #$TextBox1.BackColor = 'snow'
    $TextBox1.BackColor = 'snow'
    $TextBox1.Left = 10;
    $TextBox1.Top = 10;
    $TextBox1.Height = 40
    $TextBox1.width = 340;
    $TextBox1.Font = $Font2
    $TextBox1.add_keyup($AutoCompleteKeyupEventhandler) 
    
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
    $Panel2.Margin.All = 1
        
    #To adjust output Panel size accordint to maximum sizes, to avoid data or image getting cropped.
    $Panel2.Width = ($Panel2.Controls.width | measure -Maximum).maximum
    $Panel2.Height = ($Panel2.Controls.Height | Measure -Sum).sum + 50
    
    $Panel2.Controls.Add($AutocompleteLabel)
    
    #Define Sub Panel 3
    $Global:StatusPanel = new-object System.Windows.Forms.FlowLayoutPanel
    $StatusPanel.AutoSize = $True
    $StatusPanel.Visible = $False
    $StatusPanel.Controls.Add($ProgressBar)
    
    #Add all panels to the root Panel, so that the flow direction is Top to Down.
    $RootPanel.Controls.Add($Panel1)
    $RootPanel.Controls.Add($StatusPanel)
    $RootPanel.Controls.Add($Panel2)
    
    #Define Sub Panel 1
    $Panel1 = new-object System.Windows.Forms.FlowLayoutPanel
    $Panel1.AutoSize = $True
    
    #Define Sub Panel 2
    $Panel2 = new-object System.Windows.Forms.FlowLayoutPanel
    $Panel2.AutoSize = $True
    $Panel2.FlowDirection = 'topdown'
        
    #To adjust output Panel size accordint to maximum sizes, to avoid data or image getting cropped.
    $Panel2.Width = ($Panel2.Controls.width | measure -Maximum).maximum
    $Panel2.Height = ($Panel2.Controls.Height | Measure -Sum).sum + 50
    
    #Define Sub Panel 3
    $StatusPanel = new-object System.Windows.Forms.FlowLayoutPanel
    $StatusPanel.AutoSize = $True
    $StatusPanel.Visible = $False

#endregion Variable Definition

#region Function Definition
    
    Function Invoke-BingAutoComplete
    {
        Return (Invoke-RestMethod -Uri "http://api.bing.com/qsml.aspx?query=$($TextBox1.text)").searchsuggestion.section.item.text
    }
    
    #Function to fetch the data from Wolfram|Alpha API based on user query
    Function Invoke-WolframAlphaAPI($Global:Query)
    {
    Return (Invoke-RestMethod -Uri "http://api.wolframalpha.com/v2/query?appid=46XTUT-6T5H7K4V32&input=$($Query.Replace(' ','%20'))").queryresult
    }
    
    #Extract Results to HTML fileh
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
    
    
    #Main Funtion to Create the Basic form and its Structure.
    Function Main
    {    
    
        $Panel1.Controls.Add($TextBox1)
        $Panel1.Controls.Add($Button)
        $Panel1.Controls.Add($SaveButton)
        
    
        $Panel2.Controls.Add($AutocompleteLabel)
    
        $StatusPanel.Controls.Add($ProgressBar)
    
        $RootPanel.Controls.Add($Panel1)
        $RootPanel.Controls.Add($StatusPanel)
        $RootPanel.Controls.Add($Panel2)
        
        #Add Root Panel to the Form and display it.
        $Form.Controls.Add($RootPanel)
        [void]$Form.ShowDialog()      
    }
    
    #Function to Create the data structure for Output on Panel 3
    Function DisplayResults($Global:Result)
    {
        Try
        {
            If($Result.success -eq $True)
            {
    
                #Formula to calculate Progress bar increment each time a Sub Pod is parsed
                $Increment = (100/[int]$Result.numpods)
    
                $i=0 #Initialize ProgressBar Value 
    
                $Global:StatusLabel = New-Object Windows.forms.label
                $StatusLabel.AutoSize = $True
                $StatusLabel.Text = "$($Result.datatypes) ( "+ $("{0:N2}" -f [decimal]$Result.timing) + " Seconds )"
                $StatusLabel.Font = $Italicfont
                $StatusLabel.ForeColor = "mediumvioletred"
                $StatusPanel.Controls.Add($StatusLabel)
    
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

                            #if($s.img)
                            #{
                            #     #Create new PictureBox for all Sub POD Images
                            #     $pictureBox = new-object Windows.Forms.PictureBox
                            #     $pictureBox.Load($s.img.src)
                            #     $pictureBox.SizeMode = 'AutoSize'
                            #     $Panel2.controls.add($pictureBox)                 
                            #}
                            #Else
                            #{
                            #     #Create new Label for all Sub POD plain text
                            #     $Label = New-Object Windows.forms.label
                            #     $Label.AutoSize = $True
                            #     $Label.Text = $s.plaintext
                            #     $Panel2.Controls.Add($Label)
                            #}  
                        }
    
                #Increment the ProgressBar and display increasing values
                $i=$i+$Increment
                $ProgressBar.Value = $i
                Write-host $i
    
                }
            }
            ElseIf($Result.didyoumeans.didyoumean)
            {
                $DidYouMeans =  $Result.didyoumeans.didyoumean
                            
                $Global:DidYouMeanLabel = New-Object System.Windows.Forms.Label
                $DidYouMeanLabel.Font = $Font
                $DidYouMeanLabel.Text = "Did you mean ?"
                $DidYouMeanLabel.AutoSize = $True
                $Panel2.Controls.Add($DidYouMeanLabel)
                
                #$i=0
                #$GLobal:DidYouMeanButton = @()
                #$GLobal:DidYouMeanText = @()
    
                Foreach($DidYouMean in $DidYouMeans)
                {
                    
                    $Global:DidYouMeanButton = New-Object System.Windows.Forms.Button
                    $DidYouMeanText = $DidYouMean."#text"
                    $DidYouMeanButton.Text = "$DidYouMeanText"
                    $DidYouMeanButton.AutoSize = $True
                    $DidYouMeanButton.ForeColor = "White"
                    $DidYouMeanButton.BackColor = "Black"
                    $DidYouMeanButton.Font = $Font
                    $DidYouMeanButton.Add_Click({
                                    
                                    $Panel2.Visible = $False
                                    $Panel2.Controls.clear()
                                    $TextBox1.Text = $DidYouMeanText
                                    $DidYouMeanButton.visible = $False 
                                    $StatusPanel.Controls.clear() 
                                    $StatusPanel.controls.add($ProgressBar)
                                    $ProgressBar.value = 0
                                    $StatusPanel.Visible = $True                              
                                    DisplayResults $(Invoke-WolframAlphaAPI $DidYouMeanText)
                                    $Panel2.Visible = $True
                                    $Button.Enabled = $True
                                    $StatusPanel.controls.remove($ProgressBar)

                                  }.GetNewClosure())
    
                    $Panel2.Controls.Add($DidYouMeanButton)
                    
                }
                #$DidYouMeanButton|%{$Panel2.Controls.Add($_)}
    
                #For($i = 0;$i -le $DidYouMeans.count;$i++)
                #{
                #       $DidYouMeansArray
                #}
                #
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
        }
        catch
        {
            $Label = New-Object System.Windows.Forms.Label
            $Label.Text = "Something went wrong, Please close the window and try again"
            $Label.AutoSize = $True
            $Label.Font = $Font
            $Label.ForeColor = 'red'
            $Panel2.Controls.Add($Label)
        
        }
    }

#endregion function definition

#Calling the Function to start the tool
Main
