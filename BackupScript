###############
# Copy script #
#-------------#
# V2.1 - DZ   #
###############

###
# Changelog
###
<#
    V2.1 Changed copied dir value to actual copied dirs (e.g. robocopy always copies folder with /MIR but skips idents)
    V2.0 Updated layout and incorperated robocopy log into html layout
    V1.5 Added layout to mail template
    V1.4 Added extra try-catches and securities
    V1.3 Added robocopy log to mail template
    V1.2 Added error logging
    V1.1 Added robocopy fault messages (exit codes)
    V1.0 Basic robocopy with mail support
#>



# Clear output for testing
cls

# Global vars
$title = "Backup of YourDeviceHere : FolderNameHere"
$startTime = Get-Date -UFormat %T
$startDate = Get-Date -UFormat %Y-%m-%d

# Robocopy vars
$src = "\\Source\Folder"
$dst = "\\Destination\Folder"
$log = "\\Destination\LogFolder\Logfile.log"
$exclude = "$src\#recycle"

# Clear log file
Clear-Content $log -Filter "*.log" -Force

# Mail vars
$smtp = "YourMailServerHere"
$to = "ToMailAddress"
$from = "FromBackupServerMailAddress"
$attachment = $log
$endTime = Get-Date -UFormat %T
$endDate = Get-Date -UFormat %Y-%m-%d

###
# Copy
###

# Try to execute the robocopy
Try
{
	# Execute robocopy
    robocopy.exe "$src" "$dst" /mir /mt:96 /r:3 /w:3 /np /ts /bytes /xd $exclude /log:"$log"
	
	# Read out log results
	$result1 = (cat $log | ? { $_ -like '*Bytes :*' })
    $result2 = (cat $log | ? { $_ -like '*Dirs :*' })
    $result3 = (cat $log | ? { $_ -like '*Files :*' })

	$values1 = @()
    $values2 = @()
    $values3 = @()
	
	# Read bytes
	#$result1 | Select-String -Pattern "\d{1,}([.]\d{1,})?" -AllMatches | % { $values1 += $_.Matches.Value }
	$result1 | Select-String -Pattern "\d{1,}" -AllMatches | % { $values1 += $_.Matches.Value }
    $sizeToDivide = 1GB

	$Total = [math]::Round(($values1[0] / $sizeToDivide), 3)
	$Copied = [math]::Round(($values1[1] / $sizeToDivide), 3)
	$Skipped = [math]::Round(($values1[2] / $sizeToDivide), 3)
	$Mismach = [math]::Round(($values1[3] / $sizeToDivide), 3)
	$Failed = [math]::Round(($values1[4] / $sizeToDivide), 3)
	$Extras = [math]::Round(($values1[5] / $sizeToDivide), 3)

	# Read dirs
    $result2 | Select-String -Pattern "\d{3,}" -AllMatches | % { $values2 += $_.Matches.Value }

    $totalDirs = $values2[0]
    $copiedDirs = $values2[1]
    $skippedDirs = $values2[2]
    $mismatchDirs = $values2[3]
    $failedDirs = $values2[4]
    $extrasDirs = $values2[5]

    $actualCopiedDirs = $totalDirs - $skippedDirs

	# Read files
    $result3 | Select-String -Pattern "\d{1,}" -AllMatches | % { $values3 += $_.Matches.Value }

    $totalFiles = $values3[0]
    $copiedFiles = $values3[1]
    $skippedFiles = $values3[2]
    $mismatchFiles = $values3[3]
    $failedFiles = $values3[4]
    $extrasFiles = $values3[5]

    # No errors occurred and no files were copied.
    if($lastexitcode -eq 0)
    {
        # Clear errors
        $ErrorMessage = ""
        $FailedItem = ""

        # Set error state false (no error)
        $ErrorState = "false"

        # Set message
        $RoboMessage = "No errors occurred and no files were copied."
    }

    # One of more files were copied successfully.
    elseif($lastexitcode -eq 1)
    {
        # Clear errors
        $ErrorMessage = ""
        $FailedItem = ""

        # Set error state false (no error)
        $ErrorState = "false"

        # Set message
        $RoboMessage = "One of more files were copied successfully."
    }

    # Extra files or directories were detected.
    # Examine the log file for more information.
    elseif($lastexitcode -eq 2)
    {
        # Clear errors
        $ErrorMessage = ""
        $FailedItem = ""

        # Set error state false (no error)
        $ErrorState = "false"

        # Set message
        $RoboMessage = "Extra files or directories were detected. <br/>Examine the log file for more information."
    }

    # Some files were copied.
    # Additional files were present.
    # No failure was encountered.
    # Examine the log file for more information.
    elseif($lastexitcode -eq 3)
    {
        # Clear errors
        $ErrorMessage = ""
        $FailedItem = ""

        # Set error state false (no error)
        $ErrorState = "false"

        # Set message
        $RoboMessage = "Some files were copied. Additional files were present. No failure was encountered."
    }

    # Mismatched files or directories were detected.
    # Examine the log file for more information.
    elseif($lastexitcode -eq 4)
    {
        # Clear errors
        $ErrorMessage = "Robocopy error"
        $FailedItem = $lastexitcode

        # Set error state true (error)
        $ErrorState = "true"

        # Set message
        $RoboMessage = "Mismatched files or directories were detected. <br/>Examine the log file for more information."
    }

    # Some files were copied.
    # Some files were mismatched.
    # No failure was encountered.
    elseif($lastexitcode -eq 5)
    {
        # Clear errors
        $ErrorMessage = ""
        $FailedItem = ""

        # Set error state false (no error)
        $ErrorState = "false"

        # Set message
        $RoboMessage = "Some files were copied. Some files were mismatched. No failure was encountered."
    }

    # Additional files and mismatched files exist.
    # No files were copied and no failures were encountered.
    # This means that the files already exist in the destination directory
    elseif($lastexitcode -eq 6)
    {
        # Clear errors
        $ErrorMessage = ""
        $FailedItem = ""

        # Set error state false (no error)
        $ErrorState = "false"

        # Set message
        $RoboMessage = "Additional files and mismatched files exist. No files were copied and no failures were encountered.<br />This means that the files already exist in the destination directory"
    }

    # Files were copied, a file mismatch was present, and additional files were present.
    elseif($lastexitcode -eq 7)
    {
        # Clear errors
        $ErrorMessage = ""
        $FailedItem = ""

        # Set error state false (no error)
        $ErrorState = "false"

        # Set message
        $RoboMessage = "Files were copied, a file mismatch was present, and additional files were present."
    }

    # Some files or directories could not be copied and the retry limit was exceeded.
    elseif($lastexitcode -eq 8)
    {
        # Clear errors
        $ErrorMessage = "Robocopy error"
        $FailedItem = $lastexitcode

        # Set error state true (error)
        $ErrorState = "true"

        # Set message
        $RoboMessage = "Some files or directories could not be copied and the retry limit was exceeded."
    }

    # Robocopy did not copy any files.
    # Check the command line parameters and verify that Robocopy has enough rights to write to the destination folder.
    elseif($lastexitcode -eq 16)
    {
        # Clear errors
        $ErrorMessage = "Robocopy error"
        $FailedItem = $lastexitcode

        # Set error state true (error)
        $ErrorState = "true"

        # Set message
        $RoboMessage = "Robocopy did not copy any files. <br/>Check the command line parameters and verify that Robocopy has enough rights to write to the destination folder."
    }

    # An unknow error occurred during robocopy
    else
    {
        # Set errors
        $ErrorMessage = "An unknow error occurred"
        $FailedItem = $lastexitcode

        # Set error state true (error)
        $ErrorState = "true"

        # Set message
        $RoboMessage = "An unknow error occurred during robocopy"
    }
}

# Catch the error if it is an exception
Catch
{
    # Set errors
    $ErrorMessage = $_.Exception.Message
    $FailedItem = $_.Exception.ItemName

    # Set error state true (error)
    $ErrorState = "true"

    # Set message
    $RoboMessage = "An exception error has occurred, check the logs!"
}



###
# Send mail
###

# Time calculation
$Time1 = $startTime
$Time2 = $endTime
$TimeDiff = New-TimeSpan $Time1 $Time2

if ($TimeDiff.Seconds -lt 0)
{
	$Hrs = ($TimeDiff.Hours) + 23
	$Mins = ($TimeDiff.Minutes) + 59
	$Secs = ($TimeDiff.Seconds) + 59
}

else
{
	$Hrs = $TimeDiff.Hours
	$Mins = $TimeDiff.Minutes
	$Secs = $TimeDiff.Seconds
}

$Difference = '{0:00}:{1:00}:{2:00}' -f $Hrs,$Mins,$Secs



# If an error occured send errormessage with log
if ($ErrorState -eq "true")
{
    # Subject
    $subject = "[Failed] $title (Copied $Copied GB)"

    # Message
    $body =
    "
    <?xml version='1.0' encoding='utf-16'?>
	<!DOCTYPE html>
    <html>
	    <head>
		    <meta http-equiv='Content-Type' content='text/html; charset=utf-8'>
            
            <style>
              table
              {
                  font-family: arial, sans-serif;
                  border-collapse: collapse;
                  width: 100%;
              }

              td, th
              {
                  /*border: 1px solid #dddddd;*/
                  text-align: left;
                  padding: 10px;
              }

              tr:nth-child(even)
              {
                  #background-color: #dddddd;
              }

              #header1
              {
                  color: white;
                  text-align: left;
                  font-weight: bold;
                  font-size: 20px;
                  height: 50px;
                  vertical-align: bottom;
                  padding: 0 0 15px 15px;
                  border: none;
                  background-color: #fb9895;
              }

              #header2
              {
                  color: white;
                  text-align: center;
                  font-weight: bold;
                  font-size: 20px;
                  height: 50px;
                  vertical-align: bottom;
                  padding: 0 0 15px 15px;
                  border: none;
                  background-color: #fb9895;
              }
              
              #subheader
              {
              	margin-top: 5px;
              	font-size: 12px;
              }

              #submenu1
              {
                background-color: #f3f4f4;
              }
            </style>
            
	    </head>

	    <body>

		    <table>
			    <tr>

				    <td>
					    <table>
					
						    <tr>

							    <td id='header1'>
								    <b>$title</b>
								
								    <div id='subheader'>The backup has <b>ran</b>, please check attached log.</div>
							    </td>

							    <td id='header2'>
								    <b>Failed</b>
								
								    <div id='subheader'>
										File size copied = $Copied GB
								    </div>
							    </td>

						    </tr>

					    </table>
					
					    <table>
							
							<tr>
								<th id='submenu1' colspan='4'>Date of copy $startDate</th>
							</tr>
						
						    <tr>
							    <th width='10%'>Start time</th>
							    <td width='10%'>$startDate $startTime</td>
						    </tr>
						
						    <tr>
							    <th width='10%'>End time</th>
							    <td width='10%'>$endDate $endTime</td>
						    </tr>
						
						    <tr>
							    <th width='10%'>Duration</th>
							    <td width='10%'>$Difference</td>
						    </tr>
                            
                            <tr>
							    <th>&nbsp;</th>
						    </tr>
							
							<tr>
								<th id='submenu1' colspan='4'>Copy details</th>
							</tr>
							
							<tr>
								<th width='10%'>&nbsp;</th>
								<th width='10%'>GB</th>
								<th width='10%'>Directories</th>
								<th width='10%'>Files</th>
							</tr>

							<tr>
							    <th width='10%'>Total</th>
								<td width='10%'>$Total</td>
								<td width='10%'>$totalDirs</td>
								<td width='10%'>$totalFiles</td>
						    </tr>
							
							<tr>
							    <th width='10%'>Copied</th>
								<td width='10%'>$Copied</td>
								<td width='10%'>$actualCopiedDirs</td>
								<td width='10%'>$copiedFiles</td>
						    </tr>
							
							<tr>
							    <th width='10%'>Skipped</th>
								<td width='10%'>$Skipped</td>
								<td width='10%'>$skippedDirs</td>
								<td width='10%'>$skippedFiles</td>
						    </tr>
                            
                            <tr>
							    <th>&nbsp;</th>
						    </tr>
                            
							<tr>
								<th id='submenu1' colspan='4'>Robocopy information</th>
							</tr>
							
						    <tr>
							    <th>Source</th>
							    <td colspan='3'>$src</td>
						    </tr>
						
						    <tr>
							    <th>Destination</th>
								<td colspan='3'>$dst</td>
						    </tr>

                            <tr>
								<th>Error message</th>
								<td colspan='3'>$ErrorMessage</td>
							</tr>

                            <tr>
								<th>Errorcode</th>
								<td colspan='3'>$FailedItem</td>
							</tr>
							
							<tr>
							    <th>Robocopy message</th>
								<td colspan='3'>$RoboMessage</td>
						    </tr>

					    </table>

				    </td>
			    </tr>
		    </table>
	    </body>
    </html>
    "
}

# If an error has not occurred send the normal message with progresslog
elseif ($ErrorState -eq "false")
{
    # Subject
    $subject = "[Success] $title (Copied $Copied GB)"

    # Message
    $body = 
    "
	<?xml version='1.0' encoding='utf-16'?>
	<!DOCTYPE html>
	<html>
		<head>
			<meta http-equiv='Content-Type' content='text/html; charset=utf-8'>

			<style>
				table
                {
					font-family: arial, sans-serif;
					border-collapse: collapse;
					width: 100%;
				}

				td, th
                {
					/*border: 1px solid #dddddd;*/
					text-align: left;
					padding: 10px;
				}

				tr:nth-child(even)
                {
					#background-color: #dddddd;
				}

				#header1
				{
					color: white;
					text-align: left;
					font-weight: bold;
					font-size: 20px;
					height: 50px;
					vertical-align: bottom;
					padding: 0 0 15px 15px;
					border: none;
					background-color: #00B050;
				}

				#header2
				{
					color: white;
					text-align: center;
					font-weight: bold;
					font-size: 20px;
					height: 50px;
					vertical-align: bottom;
					padding: 0 0 15px 15px;
					border: none;
					background-color: #00B050;
				}

				#subheader
				{
					margin-top: 5px;
					font-size: 12px;
				}

				#submenu1
				{
					background-color: #f3f4f4;
				}
			</style>

		</head>

		<body>

			<table>
				<tr>
					<td>
						<table>

							<tr>

								<td id='header1'>
									<b>$title</b>
									<div id='subheader'>The backup has <b>ran</b>, please check attached log.</div>
								</td>

								<td id='header2'>
									<b>Success</b>
									<div id='subheader'>File size copied = $Copied GB</div>
								</td>

							</tr>

						</table>

						<table>

							<tr>
								<th id='submenu1' colspan='4'>Date of copy $startDate</th>
							</tr>

							<tr>
								<th width='10%'>Start time</th>
								<td width='10%'>$startDate $startTime</td>
							</tr>

							<tr>
								<th width='10%'>End time</th>
								<td width='10%'>$endDate $endTime</td>
							</tr>

							<tr>
								<th width='10%'>Duration</th>
								<td width='10%'>$Difference</td>
							</tr>

							<tr>
								<th>&nbsp;</th>
							</tr>

							<tr>
								<th id='submenu1' colspan='4'>Copy details</th>
							</tr>

							<tr>
								<th width='10%'>&nbsp;</th>
								<th width='10%'>GB</th>
								<th width='10%'>Directories</th>
								<th width='10%'>Files</th>
							</tr>

							<tr>
								<th width='10%'>Total</th>
								<td width='10%'>$Total</td>
								<td width='10%'>$totalDirs</td>
								<td width='10%'>$totalFiles</td>
							</tr>

							<tr>
								<th width='10%'>Copied</th>
								<td width='10%'>$Copied</td>
								<td width='10%'>$actualCopiedDirs</td>
								<td width='10%'>$copiedFiles</td>
							</tr>

							<tr>
								<th width='10%'>Skipped</th>
								<td width='10%'>$Skipped</td>
								<td width='10%'>$skippedDirs</td>
								<td width='10%'>$skippedFiles</td>
							</tr>

							<tr>
								<th>&nbsp;</th>
							</tr>

							<tr>
								<th id='submenu1' colspan='4'>Robocopy information</th>
							</tr>

							<tr>
								<th>Source</th>
								<td colspan='3'>$src</td>
							</tr>

							<tr>
								<th>Destination</th>
								<td colspan='3'>$dst</td>
							</tr>

							<tr>
								<th>Robocopy message</th>
								<td colspan='3'>$RoboMessage</td>
							</tr>

						</table>

					</td>
				</tr>
			</table>
		</body>
	</html>
    "
}

# If a major error occured send errormessage with log
else
{
    # Subject
    $subject = "[Failed] $title (Errorcode $FailedItem)"

    # Message
    $body =
    "
    <?xml version='1.0' encoding='utf-16'?>
	<!DOCTYPE html>
    <html>
	    <head>
		    <meta http-equiv='Content-Type' content='text/html; charset=utf-8'>
            
            <style>
              table
              {
                  font-family: arial, sans-serif;
                  border-collapse: collapse;
                  width: 100%;
              }

              td, th
              {
                  /*border: 1px solid #dddddd;*/
                  text-align: left;
                  padding: 10px;
              }

              tr:nth-child(even)
              {
                  #background-color: #dddddd;
              }

              #header1
              {
                  color: white;
                  text-align: left;
                  font-weight: bold;
                  font-size: 20px;
                  height: 50px;
                  vertical-align: bottom;
                  padding: 0 0 15px 15px;
                  border: none;
                  background-color: #fb9895;
              }

              #header2
              {
                  color: white;
                  text-align: center;
                  font-weight: bold;
                  font-size: 20px;
                  height: 50px;
                  vertical-align: bottom;
                  padding: 0 0 15px 15px;
                  border: none;
                  background-color: #fb9895;
              }
              
              #subheader
              {
              	margin-top: 5px;
              	font-size: 12px;
              }

              #submenu1
              {
                background-color: #f3f4f4;
              }
            </style>
            
	    </head>

	    <body>

		    <table>
			    <tr>

				    <td>
					    <table>
					
						    <tr>

							    <td id='header1'>
								    <b>$title</b>
								
								    <div id='subheader'>The backup has <b>failed</b>!</div>
							    </td>

							    <td id='header2'>
								    <b>Failed</b>
								
								    <div id='subheader'>
										Exitcode = $FailedItem
								    </div>
							    </td>

						    </tr>

					    </table>
					
					    <table>

							<tr>
								<th id='submenu1' colspan='4'>Robocopy information</th>
							</tr>
							
						    <tr>
							    <th>Source</th>
							    <td colspan='3'>$src</td>
						    </tr>
						
						    <tr>
							    <th>Destination</th>
								<td colspan='3'>$dst</td>
						    </tr>

                            <tr>
								<th>Error message</th>
								<td colspan='3'>$ErrorMessage</td>
							</tr>

                            <tr>
								<th>Errorcode</th>
								<td colspan='3'>$FailedItem</td>
							</tr>
							
							<tr>
							    <th>Robocopy message</th>
								<td colspan='3'>$RoboMessage</td>
						    </tr>

					    </table>

				    </td>
			    </tr>
		    </table>
	    </body>
    </html>
    "
}

# Send normal e-mail
if([string]::IsNullOrEmpty($attachment))
{
    send-MailMessage -SmtpServer $smtp -To $to -From $from -Subject $subject -Body $body -BodyAsHtml

    exit
}

else
{
    send-MailMessage -SmtpServer $smtp -To $to -From $from -Subject $subject -Body $body -Attachments $attachment -BodyAsHtml

    exit
}
