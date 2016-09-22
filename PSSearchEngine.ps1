#region Event Handlers
    
    #Event handlers
    $EventHandler =[System.EventHandler]{
    
                                    $Panel2.Visible = $False
                                    $Panel2.Controls.clear()
                                    $RelatedQueriesPanel.Controls.Clear()
                                    $StatusPanel.Controls.clear()
                                    $StatusPanel.controls.add($ProgressBar)
                                    $StatusPanel.controls.add($StatusLabel)
                                    $ProgressBar.value = 10
                                    $StatusPanel.Visible = $True
                                    $Button.Enabled = $False
                                    $StatusLabel.Text = "Computing and Fetching Results"                                
                                    DisplayResults $(Invoke-WolframAlphaAPI $TextBox1.Text)
                                    $ProgressBar.value = 20
                                    $Panel2.Visible = $True
                                    $Button.Enabled = $True
                                    $ExpanderButton.Visible = $True
                                    $ContractButton.Visible = $False
                                    $StatusPanel.controls.remove($ProgressBar)
                                    #$RelatedQueriesPanel.Controls.Clear()
                                    #$RelatedQueriesPanel.Visible = $True
                                    #$RelatedQueriesPanel.Controls.Add($ExpanderButton)
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
    }
    
    $DidYouMeanEventHandler =[System.EventHandler]{
                                    
                                    $Panel2.Visible = $False
                                    $Panel2.Controls.clear()
                                    $TextBox1.Text = $DidYouMeanText
                                    $DidYouMeanButton.visible = $False 
                                    $StatusPanel.Controls.clear()                               
                                    $StatusPanel.controls.add($StatusLabel)
                                    $ProgressBar.value = 10
                                    $StatusPanel.Visible = $True
                                    $StatusLabel.Text = "Computing and Fetching Results ..."
                                    DisplayResults $(Invoke-WolframAlphaAPI $DidYouMeanText)
                                    $ProgressBar.value = 20
                                    $Panel2.Visible = $True
                                    $Button.Enabled = $True
                                    $StatusPanel.controls.remove($ProgressBar)
                                    #$RelatedQueriesPanel.Controls.Clear()
                                    #$RelatedQueriesPanel.Controls.Add($ExpanderButton)
                                  }.GetNewClosure()
    
    $AutoCompleteKeyupEventhandler =  [System.Windows.Forms.KeyEventHandler]{                           
                                    
                                    $Panel2.Controls.clear()
                                    $RelatedQueriesPanel.Controls.Clear()
                                    $StatusPanel.Controls.clear()
                                    $StrWithLineBreaks=@()
                                    $Data = Invoke-BingAutoComplete
                                    $Data | %{$StrWithLineBreaks+=$_+';'}
                                    
                                    #typecasted String[] tp String
                                    $AutocompleteLabel.text=[string]($StrWithLineBreaks) -replace "; ","`n" -replace ";",""
                                    $Panel2.Controls.Add($AutocompleteLabel)
    }

    $RelatedQueryExpanderEventHandler =[System.EventHandler]{
    
                                    If(-not($RelatedQueriesPanel.Controls|? text -eq '+'))
                                    {
                                        $RelatedQueriesPanel.Controls.Add($ExpanderButton)
                                    }
                                    
                                    $RelatedQueriesPanel.Controls.add($ContractButton)
                                    
                                    #Add related queries as a button onto related queries panel
                                    Foreach($Rq in $RelatedQueries)
                                    {
                                    
                                        $Global:RelatedQueryButton = New-Object System.Windows.Forms.Button
                                        $RelatedQueryButton.Text = (Get-Culture).TextInfo.ToTitleCase("$Rq")
                                        $RelatedQueryButton.AutoSize = $True
                                        $RelatedQueryButton.BackColor = "White"
                                        $RelatedQueryButton.ForeColor = "Black"
                                        $RelatedQueryButton.Font = $RegularFontBig
                                        $RelatedQueryButton.FlatStyle = 'Flat'
                                        $RelatedQueryButton.FlatAppearance.BorderColor = 'Black'
                                        $RelatedQueryButton.FlatAppearance.BorderSize = 1
                                        $RelatedQueryButton.FlatAppearance.MouseOverBackColor = 'lightyellow'
                                        $RelatedQueryButton.AutoSizeMode = 'GrowAndShrink'
                                    
                                        $RelatedQueryButton.Add_Click({
                                    
                                                                        $Panel2.Visible = $False
                                                                        $Panel2.Controls.clear()
                                                                        $TextBox1.Text = $Rq
                                                                        $StatusPanel.Controls.clear() 
                                                                        $RelatedQueriesPanel.Controls.clear()
                                                                        $StatusPanel.controls.add($ProgressBar)
                                                                        $StatusPanel.controls.add($StatusLabel)
                                                                        $ProgressBar.value = 10
                                                                        $StatusPanel.Visible = $True
                                                                        $StatusLabel.Text = "Computing Fetching Results ..."
                                                                        DisplayResults $(Invoke-WolframAlphaAPI $Rq)
                                                                        $ProgressBar.value = 20
                                                                        $ExpanderButton.Visible = $true
                                                                        $ContractButton.Visible = $False
                                                                        $Panel2.Visible = $True
                                                                        $Button.Enabled = $True
                                                                        $StatusPanel.controls.remove($ProgressBar)
                                                                        #$RelatedQueriesPanel.Controls.Clear()
                                                                        #$RelatedQueriesPanel.Controls.Add($ExpanderButton)
                                    
                                        }.GetNewClosure())
                                    
                                        $RelatedQueriesPanel.controls.add($RelatedQueryButton)
                                    }
                                                                        
                                    $ContractButton.Visible = $True  
                                    $ExpanderButton.Visible = $False
    }

    $RelatedQueryContractEventHandler = [System.EventHandler]{

                                                             $RelatedQueriesPanel.Controls.Clear()
                                                             $RelatedQueriesPanel.Controls.add($RelatedQueryLabel)
                                                             $RelatedQueriesPanel.Controls.add($ExpanderButton)
                                                             $ExpanderButton.Visible = $True
                                                             #$ContractButton.Visible = $False
    }

#endregion Event Handlers

#region Variable Definition

    #Calling the Assemblies
    [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
    #[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")
    
    #Define Text Font object
    $FontFamily = "Lucida sans"

    $ItalicFont = New-Object System.Drawing.Font($FontFamily,8,[System.Drawing.FontStyle]::Italic) 
    $ItalicFontBig = New-Object System.Drawing.Font($FontFamily,10,[System.Drawing.FontStyle]::Italic) 
    $RegularFont = New-Object System.Drawing.Font($FontFamily,10,[System.Drawing.FontStyle]::Regular) 
    $RegularFontBig = New-Object System.Drawing.Font($FontFamily,11,[System.Drawing.FontStyle]::Regular) 
    $BoldFont = New-Object System.Drawing.Font($FontFamily,11,[System.Drawing.FontStyle]::bold) 
    $BoldFontBig = New-Object System.Drawing.Font($FontFamily,13,[System.Drawing.FontStyle]::bold) 
    
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
    $Form.AutoScroll = $true
    
    #Define the Base Panel on which we'll add 4 sub panels
    $RootPanel = new-object System.Windows.Forms.FlowLayoutPanel
    $RootPanel.AutoSize = $True
    $RootPanel.FlowDirection = 'topdown'
    
    #Define Panel 1
    $Panel1 = new-object System.Windows.Forms.FlowLayoutPanel
    $Panel1.AutoSize = $True

        #region Panel1 items
            
            #Define TextBox1 for input
            $TextBox1 = New-Object “System.Windows.Forms.RichTextBox”;
            $TextBox1.BorderStyle = 'fixed3d'
            $TextBox1.BackColor = 'snow'
            $TextBox1.Multiline = $true
            $TextBox1.Left = 10;
            $TextBox1.Top = 10;
            $TextBox1.Height = 40
            $TextBox1.width = 340;
            $TextBox1.Font = $BoldFontBig
            $TextBox1.add_keyup($AutoCompleteKeyupEventhandler) 
            
            #Define Search Button
            $Button = New-Object System.Windows.Forms.Button
            $Button.Text = "Search"
            $Button.Font = $BoldFontBig
            $Button.Height = 40
            $Button.Add_Click($EventHandler)
            
            #Define Save Button
            $SaveButton = New-Object System.Windows.Forms.Button
            $SaveButton.text = "Save"    
            $SaveButton.Font = $BoldFontBig
            $SaveButton.Height =  40
            $SaveButton.Add_Click($SaveEventHandler)

        #endregion Panel1 items

    #Define Status Panel
    $Global:StatusPanel = new-object System.Windows.Forms.FlowLayoutPanel
    $StatusPanel.AutoSize = $True
    $StatusPanel.Visible = $False
    $StatusPanel.FlowDirection = 'topdown'


        #region StatusPanel Items
            
            #Define the Progress Bar
            $Global:ProgressBar = New-Object System.Windows.Forms.ProgressBar
            $ProgressBar.Maximum = 100
            $ProgressBar.Minimum = 0
            $ProgressBar.Height = 10
            $ProgressBar.Width = 500
            $ProgressBar.BackColor = 'Blue'
            $ProgressBar.Style = 'Blocks'
            $ProgressBar.Visible = $true
            $ProgressBar.Enabled
            $ProgressBar.Value = 5
            #$progressbar.Text

            $StatusLabel = New-Object Windows.forms.label
            $StatusLabel.AutoSize = $True
            $StatusLabel.Visible = $True
            $StatusLabel.Font = $Italicfont
            $StatusLabel.ForeColor = "mediumvioletred"
                    
        #endregion StatusPanel Items

    #Define Related Queries Panel
    $RelatedQueriesPanel = new-object System.Windows.Forms.FlowLayoutPanel
    $RelatedQueriesPanel.AutoSize = $True
    $RelatedQueriesPanel.Visible = $True
    $RelatedQueriesPanel.FlowDirection = 'TopDown'

        #region RelatedQueriesPanel Items
                
                $RelatedQueryLabel =  New-Object System.Windows.Forms.Label
                $RelatedQueryLabel.Font = $RegularFont
                $RelatedQueryLabel.ForeColor = "navy"
                $RelatedQueryLabel.AutoSize = $True

                $ExpanderButton = New-Object System.Windows.Forms.Button
                $ExpanderButton.Text = "+"
                $ExpanderButton.TextAlign = 'middlecenter'
                $ExpanderButton.Font = $BoldFont
                $ExpanderButton.Width = 25
                $ExpanderButton.Height = 25
                $ExpanderButton.BackColor = 'Black'
                $ExpanderButton.ForeColor = 'White'
                $ExpanderButton.FlatStyle = 'flat'
                $ExpanderButton.FlatAppearance.BorderColor = 'Black'
                $ExpanderButton.FlatAppearance.BorderSize = 1
                $ExpanderButton.FlatAppearance.MouseOverBackColor = 'gray'
                $ExpanderButton.add_click($RelatedQueryExpanderEventHandler)

                
                $ContractButton = New-Object System.Windows.Forms.Button
                $ContractButton.Visible = $False
                $ContractButton.Text = "-"
                $ContractButton.TextAlign = 'middlecenter'
                $ContractButton.Font = $BoldFont
                $ContractButton.width = 25
                $ContractButton.Height = 25
                $ContractButton.BackColor = 'Black'
                $ContractButton.ForeColor = 'White'
                $ContractButton.FlatStyle = 'flat'
                $ContractButton.FlatAppearance.BorderColor = 'Black'
                $ContractButton.FlatAppearance.BorderSize = 1
                $ContractButton.FlatAppearance.MouseOverBackColor = 'gray'
                $ContractButton.add_click($RelatedQueryContractEventHandler)
        
        #endregion RelatedQueriespanelItem
        
    #Define Panel 2
    $Panel2 = new-object System.Windows.Forms.FlowLayoutPanel
    $Panel2.AutoSize = $True
    $Panel2.FlowDirection = 'topdown'
    $Panel2.Margin.All = 1    
    $Panel2.Width = ($Panel2.Controls.width | measure -Maximum).maximum #To adjust output Panel size accordint to maximum sizes, to avoid data or image getting cropped.
    $Panel2.Height = ($Panel2.Controls.Height | Measure -Sum).sum + 50

        #region Panel2 Items
            
                $AutocompleteLabel = New-Object System.Windows.Forms.Label
                $AutocompleteLabel.AutoSize = $True
                $AutocompleteLabel.Font = $RegularFont

        #endregion Panel2 Items
    
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
            
        
            "<h3>$($p.title.toupper())</h3>"
        
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

        #Add all panels to the root Panel, so that the flow direction is Top to Down.    
        $RootPanel.Controls.Add($Panel1)
        $RootPanel.Controls.Add($StatusPanel)
        $RootPanel.Controls.Add($RelatedQueriesPanel)
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
                $StatusPanel.Controls.Add($StatusLabel)
                $ProgressBar.Value = 30
                $ProgressBar.Refresh()

                $StatusLabel.Text = "Loading related queries"

                #Fetch related queries 
                $Global:RelatedQueries =  (Invoke-RestMethod -Uri $Result.related).relatedqueries.relatedquery

                #If related queries exist
                if($result.related -and $RelatedQueries)
                {
                    $RelatedQueryLabel.Text = "Found $($RelatedQueries.count) related queries click '+' below to expand"

                    #Expand/Contract functionality for Related queries
                    $RelatedQueriesPanel.Controls.Add($ExpanderButton)                
                    $RelatedQueriesPanel.Controls.Add($ContractButton)
                    $RelatedQueriesPanel.Controls.Add($RelatedQueryLabel)
                }
                
                 
                #Formula to calculate Progress bar increment each time a Sub Pod is parsed
                $Increment = (50/[int]$Result.numpods)
    
                $i=50 #Initialize ProgressBar Value 
                $StatusLabel.Text = "Generating Output"
    
                $DataType= $($Result.datatypes)
                $Timetaken =  $("{0:N2}" -f [decimal]$Result.timing)    
                
    
                If($Result.warnings.spellcheck.text)
                {
                    $spellcheck = $Result.warnings.spellcheck.text

                    if($spellcheck.count -gt 1)
                    {
                        $spellcheck = [string]$spellcheck -replace "`" Interpreting","`" AND Interpreting"
                    }

                    $WarningLabel = New-Object Windows.forms.label
                    $WarningLabel.Text = "Warning : $spellcheck"
                    $WarningLabel.Font = $RegularFont
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
                    $LabelTitle.Font = $BoldFontBig
                    $Panel2.Controls.Add($LabelTitle)

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
    
                }

                If($DataType)
                {
                    $StatusLabel.Text = "$DataType ( "+ $Timetaken + " Seconds )"
                }
                else
                {
                    $StatusLabel.Text = "Time : $Timetaken Seconds"
                }

            }
            ElseIf($Result.didyoumeans.didyoumean)
            {
                $StatusPanel.Visible = $False
                $DidYouMeans =  $Result.didyoumeans.didyoumean
                            
                $Global:DidYouMeanLabel = New-Object System.Windows.Forms.Label
                $DidYouMeanLabel.Font = $BoldFont
                $DidYouMeanLabel.Text = "Did you mean ?"
                $DidYouMeanLabel.AutoSize = $True
                
                $Panel2.Controls.Add($DidYouMeanLabel)

                Foreach($DidYouMean in $DidYouMeans)
                {
                    
                    $Global:DidYouMeanButton = New-Object System.Windows.Forms.Button
                    $DidYouMeanText = (Get-Culture).TextInfo.ToTitleCase("$($DidYouMean."#text")")
                    $DidYouMeanButton.Text = "$DidYouMeanText";
                    $DidYouMeanButton.AutoSize = $True
                    $DidYouMeanButton.BackColor = "White"
                    $DidYouMeanButton.ForeColor = "Black"
                    $DidYouMeanButton.Font = $RegularFontBig
                    $DidYouMeanButton.FlatStyle = 'Flat'
                    $DidYouMeanButton.FlatAppearance.BorderColor = 'Black'
                    $DidYouMeanButton.FlatAppearance.BorderSize = 1
                    $DidYouMeanButton.FlatAppearance.MouseOverBackColor = 'lightyellow'
                    $DidYouMeanButton.AutoSizeMode = 'GrowAndShrink'

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

            }
            ElseIf($Result.tips.tip)
            {
                    $Tips =  $Result.Tips.Tip
                
                    Foreach($Tip in $Tips)
                    {
                        
                        $Global:TipsLabel = New-Object System.Windows.Forms.Label
                        $TipsLabel.Font = $ItalicFontBig
                        $TipsLabel.Text = "TIP : $($tip.Text)"
                        $TipsLabel.AutoSize = $True
                        $TipsLabel.ForeColor = 'Navy'
                        $StatusPanel.Visible = $False
                        $Panel2.Controls.Add($tipsLabel)
                    }
            }
            else
            {
                $Label = New-Object System.Windows.Forms.Label
                $Label.Text = "No Results Found."
                $Label.AutoSize = $True
                $Label.Font = $ItalicFontBig
                $Label.ForeColor = 'Navy'
                $StatusPanel.Visible = $False
                $Panel2.Controls.Add($Label)
            }

        }
        catch
        {
            $Label = New-Object System.Windows.Forms.Label
            $Label.Text = "Something went wrong, Please close the window and try again"
            $Label.AutoSize = $True
            $Label.Font = $ItalicFontBig
            $Label.ForeColor = 'red'
            $StatusPanel.Visible = $False
            $Panel2.Controls.Add($Label)
        
        }
    }

#endregion function definition

#Calling the Function to start the tool
Main
