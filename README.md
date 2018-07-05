<!-- Visual Studio Code: For a more comfortable reading experience, use the key combination Ctrl + Shift + V
     Visual Studio Code: To crop the tailing end space characters out, please use the key combination Ctrl + A Ctrl + K Ctrl + X (Formerly Ctrl + Shift + X)
     Visual Studio Code: To improve the formatting of HTML code, press Shift + Alt + F and the selected area will be reformatted in a html file.
     Visual Studio Code shortcuts: http://code.visualstudio.com/docs/customization/keybindings (or https://aka.ms/vscodekeybindings)
     Visual Studio Code shortcut PDF (Windows): https://code.visualstudio.com/shortcuts/keyboard-shortcuts-windows.pdf


   _____      _          ______      _  __ _____        _
  / ____|    | |        |  ____|    (_)/ _|  __ \      | |
 | |  __  ___| |_ ______| |__  __  ___| |_| |  | | __ _| |_ __ _
 | | |_ |/ _ \ __|______|  __| \ \/ / |  _| |  | |/ _` | __/ _` |
 | |__| |  __/ |_       | |____ >  <| | | | |__| | (_| | || (_| |
  \_____|\___|\__|      |______/_/\_\_|_| |_____/ \__,_|\__\__,_|                                   -->


## Get-ExifData.ps1

<table>
    <tr>
        <td style="padding:6px"><strong>OS:</strong></td>
        <td colspan="2" style="padding:6px">Windows</td>
    </tr>
    <tr>
        <td style="padding:6px"><strong>Type:</strong></td>
        <td colspan="2" style="padding:6px">A Windows PowerShell script</td>
    </tr>
    <tr>
        <td style="padding:6px"><strong>Language:</strong></td>
        <td colspan="2" style="padding:6px">Windows PowerShell</td>
    </tr>
    <tr>
        <td style="padding:6px"><strong>Description:</strong></td>
        <td colspan="2" style="padding:6px">
            <p>
                Get-ExifData reads digital image files and tries to retrieve EXIF data from them and write that info to a CSV-file (<code>exif_log.csv</code>). The console displays rudimentary info about the gathering process, a reduced list is displayed in a pop-up window (<code>Out-GridView</code>, about 30 categories) and the CSV-file is written/updated with over 350 categories, including the GPS tags.</p>
            <p>
                The list of image files to be read is constructed in the command launching Get-ExifData by adding a full path of a folder (after <code>-Path</code> parameter) or by adding a full path of individual files (after <code>-File</code> parameter, multiple entries separated with a comma). The search for image files may also be done recursively by adding the <code>-Recurse</code> parameter the command launching Get-ExifData. If <code>-Path</code> and <code>-File</code> parameters are not defined, Get-ExifData reads non-recursively the image files, which reside in the "<code>$($env:USERPROFILE)\Pictures</code>" folder.</p>
            <p>
                By default the CSV-file (<code>exif_log.csv</code>) is created into the User's own picture folder "<code>$($env:USERPROFILE)\Pictures</code>" but the default CSV-file destination may be changed with the <code>-Output</code> parameter. Shall the CSV-file already exist, Get-ExifData tries to add new info to the bottom of the CSV-file rather than overwrite the CSV-file. If the user wishes not to create any logs (<code>exif_log.csv</code>) or update any existing (<code>exif_log.csv</code>) files, the <code>-SuppressLog</code> parameter may be added to the command launching Get-ExifData.</p>
            <p>
                The other available parameters (<code>-Force</code>, <code>-Open</code> and <code>-Audio</code>) are discussed in greater detail below. Please note, that if any of the individual parameter values include space characters, the individual value should be enclosed in quotation marks (single or double), so that PowerShell can interpret the command correctly.</p>
        </td>
    </tr>
    <tr>
        <td style="padding:6px"><strong>Homepage:</strong></td>
        <td colspan="2" style="padding:6px"><a href="https://github.com/auberginehill/get-exif-data">https://github.com/auberginehill/get-exif-data</a>
            <br />Short URL: <a href="http://tinyurl.com/ycbhtpba">http://tinyurl.com/ycbhtpba</a></td>
    </tr>
    <tr>
        <td style="padding:6px"><strong>Version:</strong></td>
        <td colspan="2" style="padding:6px">1.0</td>
    </tr>
    <tr>
        <td rowspan="7" style="padding:6px"><strong>Sources:</strong></td>
        <td style="padding:6px">Emojis:</td>
        <td style="padding:6px"><a href="https://github.com/auberginehill/emoji-table">Emoji Table</a></td>
    </tr>
    <tr>
        <td style="padding:6px">clayman2:</td>
        <td style="padding:6px"><a href="http://powershell.com/cs/media/p/7476.aspx">Disk Space</a> (or one of the <a href="http://web.archive.org/web/20120304222258/http://powershell.com/cs/media/p/7476.aspx">archive.org versions</a>)</td>
    </tr>
    <tr>
        <td style="padding:6px">Franck Richard:</td>
        <td style="padding:6px"><a href="http://franckrichard.blogspot.com/2011/04/2011-scripting-games-advanced-event-8.html">Use PowerShell to Remove Metadata and Resize Images</a></td>
    </tr>
    <tr>
        <td style="padding:6px">lamaar75:</td>
        <td style="padding:6px"><a href="http://powershell.com/cs/forums/t/9685.aspx">Creating a Menu</a> (or one of the <a href="https://web.archive.org/web/20150910111758/http://powershell.com/cs/forums/t/9685.aspx">archive.org versions</a>)</td>
    </tr>
    <tr>
        <td style="padding:6px">Twon of An:</td>
        <td style="padding:6px"><a href="https://community.spiceworks.com/scripts/show/2263-get-the-sha1-sha256-sha384-sha512-md5-or-ripemd160-hash-of-a-file">Get the SHA1,SHA256,SHA384,SHA512,MD5 or RIPEMD160 hash of a file</a></td>
    </tr>
    <tr>
        <td style="padding:6px">Gisli:</td>
        <td style="padding:6px"><a href="http://stackoverflow.com/questions/8711564/unable-to-read-an-open-file-with-binary-reader">Unable to read an open file with binary reader</a></td>
    </tr>
    <tr>
        <td style="padding:6px">Fred:</td>
        <td style="padding:6px"><a href="https://social.technet.microsoft.com/Forums/scriptcenter/en-US/76ae6430-4993-4422-aa97-8f8ec3ca4e87/selectobject-where?forum=winserverpowershell">select-object | where</a></td>
    </tr>
    <tr>
        <td style="padding:6px"><strong>Downloads:</strong></td>
        <td colspan="2" style="padding:6px">For instance <a href="https://raw.githubusercontent.com/auberginehill/get-exif-data/master/Get-ExifData.ps1">Get-ExifData.ps1</a>. Or <a href="https://github.com/auberginehill/get-exif-data/archive/master.zip">everything as a .zip-file</a>.</td>
    </tr>
</table>




### Screenshot

<img class="screenshot" title="screenshot" alt="screenshot" height="100%" width="100%" src="https://raw.githubusercontent.com/auberginehill/get-exif-data/master/Get-ExifData.jpg">




### Parameters

<table>
    <tr>
        <th>:triangular_ruler:</th>
        <td style="padding:6px">
            <ul>
                <li>
                    <h5>Parameter <code>-Path</code></h5>
                    <p>with aliases <code>-Directory</code>, <code>-DirectoryPath</code>, <code>-Folder</code> and <code>-FolderPath</code>.  Specifies the primary folder, from which the image files are checked for their EXIF data. The default <code>-Path</code> parameter is "<code>$($env:USERPROFILE)\Pictures</code>", which will be used, if any value for the <code>-Path</code> or the <code>-File</code> parameters is not included in the command launching Get-ExifData.</p>
                    <p>The value for the <code>-Path</code> parameter should be a valid file system path pointing to a directory (a full path of a folder such as <code>C:\Users\Dropbox\</code>). Furthermore, if the path includes space characters, please enclose the path in quotation marks (single or double). Multiple entries may be entered, if they are separated with a comma.</p>
                </li>
            </ul>
        </td>
    </tr>
    <tr>
        <th></th>
        <td style="padding:6px">
            <ul>
                <p>
                    <li>
                        <h5>Parameter <code>-File</code></h5>
                        <p>with aliases <code>-SourceFile</code>, <code>-FilePath</code> and <code>-Files</code>. Specifies, which image files are checked for their EXIF data. The value for the <code>-File</code> parameter should be a valid full file system path pointing to a file (with a full path name of a folder such as <code>C:\Windows\explorer.exe</code>). Furthermore, if the path includes space characters, please enclose the path in quotation marks (single or double). Multiple entries may be entered, if they are separated with a comma.</p>
                    </li>
                </p>
                <p>
                    <li>
                        <h5>Parameter <code>-Output</code></h5>
                        <p>with aliases <code>-OutputFolder</code> and <code>-LogFileFolder</code>. Defines the folder/directory, where the CSV-file is created or updated.  The default <code>-Output</code> parameter is "<code>$($env:USERPROFILE)\Pictures</code>", which will be used, if any value for the <code>-Output</code> is not included in the command launching Get-ExifData.</p>
                        <p>The value for the <code>-Output</code> parameter should be a valid file system path pointing to a directory (a full path of a folder such as <code>C:\Users\Dropbox\</code>). Furthermore, if the path includes space characters, please enclose the path in quotation marks (single or double).</p>
                        <p>The log file file name (<code>exif_log.csv</code>) is defined on row 78 with <code>$log_filename</code> variable and is thus "hard coded" into the script. The produced log file is UTF-8 encoded CSV-file with semi-colon as the separator.</p>
                    </li>
                </p>
                <p>
                    <li>
                        <h5>Parameter <code>-Recurse</code></h5>
                        <p>The search for image files is done recursively, i.e. if a folder/directory is found, all the subsequent subfolders and the image files that reside within those subfolders (and in the subfolders of the subfolders' subfolders, and their subfolders and so forth...) are included in the EXIF data gathering process. Please note, that with great many image files, Get-ExifData may take some time to process each and every file.</p>
                    </li>
                </p>
                <p>
                    <li>
                        <h5>Parameter <code>-SuppressLog</code></h5>
                        <p>with aliases <code>-Silent</code>, <code>-Suppress</code>, <code>-NoLog</code> and <code>-DoNotCreateALog</code>. By adding <code>-SuppressLog</code> to the command launching Get-ExifData, the CSV-file (<code>exif_log.csv</code>) is not created, touched nor updated.</p>
                    </li>
                </p>
                <p>
                    <li>
                        <h5>Parameter <code>-Open</code></h5>
                        <p>If the <code>-Open</code> parameter is used in the command launching Get-ExifData and new EXIF data is found, the CSV-file destination folder (which is defined with the <code>-Output</code> parameter) is opened in the File Manager.</p>
                    </li>
                </p>
                <p>
                    <li>
                        <h5>Parameter <code>-Force</code></h5>
                        <p>The <code>-Force</code> parameter affects the behaviour of Get-ExifData in two ways. If the <code>-Force</code> parameter is used with the...
                            <ol>
                                <li><code>-Output</code> parameter, the CSV-file destination folder (defined with the <code>-Output</code> parameter) is created, without asking any further confirmations from the end-user. The new folder is created with the command <code>New-Item "$Output" -ItemType Directory -Force</code> which may not be powerfull enough to create a new folder inside any arbitrary (system) folder. The Get-ExifData may gain additional rights, if it's run in an elevated PowerShell window (but for the most cases that is not needed at all).</li>
                                <li><code>-Open</code> parameter, the CSV-file destination folder (defined with the <code>-Output</code> parameter) is opened regardless whether any new EXIF data was found or not.</li>
                            </ol>
                        </p>
                    </li>
                </p>
                <p>
                    <li>
                        <h5>Parameter <code>-Audio</code></h5>
                        <p>If the <code>-Audio</code> parameter is used in the command launching Get-ExifData and new EXIF data is found, an audible beep will occur.</p>
                    </li>
                </p>
            </ul>
        </td>
    </tr>
</table>




### Outputs

<table>
    <tr>
        <th>:arrow_right:</th>
        <td style="padding:6px">
            <ul>
                <li>Displays a summary of the actions in console. Displays a reduced EXIF data list in a pop-up window (<code>Out-GridView</code>). Writes or updates a CSV log file at the path defined with the <code>-Output</code> parameter, if the <code>-SuppressLog</code> parameter is not used.</li>
            </ul>
        </td>
    </tr>
    <tr>
        <th></th>
        <td style="padding:6px">
            <ul>
                <ol>
                    <p>Default values:</p>
                    <p>
                        <table>
                            <tr>
                                <td style="padding:6px"><strong>Path</strong></td>
                                <td style="padding:6px"><strong>Parameter</strong></td>
                                <td style="padding:6px"><strong>Type</strong></td>
                            </tr>
                            <tr>
                                <td style="padding:6px"><code>"$($env:USERPROFILE)\Pictures\exif_log.csv"</code></td>
                                <td style="padding:6px"><code>-Output</code> (concerning the folder)</td>
                                <td style="padding:6px">CSV log file containing EXIF data</td>
                            </tr>
                            <tr>
                                <td style="padding:6px"><code>"$($env:USERPROFILE)\Pictures"</code></td>
                                <td style="padding:6px"><code>-Path</code></td>
                                <td style="padding:6px">The folder for searching the image files for their EXIF data, if no <code>-Path</code> or <code>-File</code> parameter is used (a non-recursive search).</td>
                            </tr>
                        </table>
                    </p>
                </ol>
            </ul>
        </td>
    </tr>
</table>




### Notes

<table>
    <tr>
        <th>:warning:</th>
        <td style="padding:6px">
            <ul>
                <li>Please note that all the parameters can be used in one get EXIF data command, and that each of the parameters can be "tab completed" before typing them fully (by pressing the <code>[tab]</code> key).</li>
            </ul>
        </td>
    </tr>
</table>




### Examples

<table>
    <tr>
        <th>:book:</th>
        <td style="padding:6px">To open this code in Windows PowerShell, for instance:</td>
   </tr>
   <tr>
        <th></th>
        <td style="padding:6px">
            <ol>
                <p>
                    <li><code>./Get-ExifData.ps1</code><br />
                    Runs the script. Please notice to insert <code>./</code> or <code>.\</code> before the script name. Tries to read the EXIF data from image files that reside in the "<code>$($env:USERPROFILE)\Pictures\</code>" folder, since no values for the <code>-Path</code> or <code>-File</code> parameters were defined. Saves or updates the CSV log file (<code>exif_log.csv</code>) at the default <code>-Output</code> folder (<code>"$($env:USERPROFILE)\Pictures\exif_log.csv"</code>) - a file that contains all the gathered EXIF info columns/data types. A pop-up window listing a partial list of the EXIF info will open, if image files were read. The console will show rudimentary stats about the EXIF data gathering procedure.</li>
                </p>
                <p>
                    <li><code>help ./Get-ExifData -Full</code><br />
                    Displays the help file.</li>
                </p>
                <p>
                    <li><code>.\Get-ExifData.ps1 -Path "C:\Users\Dropbox\" -Output "C:\Users\Dropbox\dc01" -Audio -Open -Recurse -Force</code><br />
                    Runs the script and tries to recursively search for image files at "<code>C:\Users\Dropbox\</code>" and read the EXIF info of the found image files and either (1) update or create the CSV log file (<code>exif_log.csv</code>) at the <code>C:\Users\Dropbox\dc01</code> folder if the folder exists or (2) create the C:\Users\Dropbox\dc01 folder (and exif_log.csv) without asking any further questions, if the <code>-Output</code> directory doesn't exist (since the <code>-Force</code> was used). Also, since the <code>-Force</code> and <code>-Open</code> parameters were used, the default File Manager will be opened at <code>C:\Users\Dropbox\dc01</code> regardless whether any image files were read or not. Furthermore, if new image files were indeed read, an audible beep will occur. Also, a pop-up window listing a partial list of the EXIF info will open, if image files were read, and the console will show rudimentary stats about the EXIF data gathering procedure.</li>
                </p>
                <p>
                    <li><p><code>Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope LocalMachine</code><br />
                    This command is altering the Windows PowerShell rights to enable script execution in the default (<code>LocalMachine</code>) scope, and defines the conditions under which Windows PowerShell loads configuration files and runs scripts in general. In Windows Vista and later versions of Windows, for running commands that change the execution policy of the <code>LocalMachine</code> scope, Windows PowerShell has to be run with elevated rights (<dfn>Run as Administrator</dfn>). The default policy of the default (<code>LocalMachine</code>) scope is "<code>Restricted</code>", and a command "<code>Set-ExecutionPolicy Restricted</code>" will "<dfn>undo</dfn>" the changes made with the original example above (had the policy not been changed before...). Execution policies for the local computer (<code>LocalMachine</code>) and for the current user (<code>CurrentUser</code>) are stored in the registry (at for instance the <code>HKLM:\Software\Policies\Microsoft\Windows\PowerShell\ExecutionPolicy</code> key), and remain effective until they are changed again. The execution policy for a particular session (<code>Process</code>) is stored only in memory, and is discarded when the session is closed.</p>
                        <p>Parameters:
                            <ul>
                                <table>
                                    <tr>
                                        <td style="padding:6px"><code>Restricted</code></td>
                                        <td colspan="2" style="padding:6px">Does not load configuration files or run scripts, but permits individual commands. <code>Restricted</code> is the default execution policy.</td>
                                    </tr>
                                    <tr>
                                        <td style="padding:6px"><code>AllSigned</code></td>
                                        <td colspan="2" style="padding:6px">Scripts can run. Requires that all scripts and configuration files be signed by a trusted publisher, including the scripts that have been written on the local computer. Risks running signed, but malicious, scripts.</td>
                                    </tr>
                                    <tr>
                                        <td style="padding:6px"><code>RemoteSigned</code></td>
                                        <td colspan="2" style="padding:6px">Requires a digital signature from a trusted publisher on scripts and configuration files that are downloaded from the Internet (including e-mail and instant messaging programs). Does not require digital signatures on scripts that have been written on the local computer. Permits running unsigned scripts that are downloaded from the Internet, if the scripts are unblocked by using the <code>Unblock-File</code> cmdlet. Risks running unsigned scripts from sources other than the Internet and signed, but malicious, scripts.</td>
                                    </tr>
                                    <tr>
                                        <td style="padding:6px"><code>Unrestricted</code></td>
                                        <td colspan="2" style="padding:6px">Loads all configuration files and runs all scripts. Warns the user before running scripts and configuration files that are downloaded from the Internet. Not only risks, but actually permits, eventually, running any unsigned scripts from any source. Risks running malicious scripts.</td>
                                    </tr>
                                    <tr>
                                        <td style="padding:6px"><code>Bypass</code></td>
                                        <td colspan="2" style="padding:6px">Nothing is blocked and there are no warnings or prompts. Not only risks, but actually permits running any unsigned scripts from any source. Risks running malicious scripts.</td>
                                    </tr>
                                    <tr>
                                        <td style="padding:6px"><code>Undefined</code></td>
                                        <td colspan="2" style="padding:6px">Removes the currently assigned execution policy from the current scope. If the execution policy in all scopes is set to <code>Undefined</code>, the effective execution policy is <code>Restricted</code>, which is the default execution policy. This parameter will not alter or remove the ("<dfn>master</dfn>") execution policy that is set with a Group Policy setting.</td>
                                    </tr>
                                    <tr>
                                        <td style="padding:6px; border-top-width:1px; border-top-style:solid;"><span style="font-size: 95%">Notes:</span></td>
                                        <td colspan="2" style="padding:6px">
                                            <ul>
                                                <li><span style="font-size: 95%">Please note that the Group Policy setting "<code>Turn on Script Execution</code>" overrides the execution policies set in Windows PowerShell in all scopes. To find this ("<dfn>master</dfn>") setting, please, for example, open the Local Group Policy Editor (<code>gpedit.msc</code>) and navigate to Computer Configuration → Administrative Templates → Windows Components → Windows PowerShell.</span></li>
                                            </ul>
                                        </td>
                                    </tr>
                                    <tr>
                                        <th></th>
                                        <td colspan="2" style="padding:6px">
                                            <ul>
                                                <li><span style="font-size: 95%">The Local Group Policy Editor (<code>gpedit.msc</code>) is not available in any Home or Starter edition of Windows.</span></li>
                                                <ol>
                                                    <p>
                                                        <table>
                                                            <tr>
                                                                <td style="padding:6px; font-size: 85%"><strong>Group Policy Setting</strong> "<code>Turn&nbsp;on&nbsp;Script&nbsp;Execution</code>"</td>
                                                                <td style="padding:6px; font-size: 85%"><strong>PowerShell Equivalent</strong> (concerning all scopes)</td>
                                                            </tr>
                                                            <tr>
                                                                <td style="padding:6px; font-size: 85%"><code>Not configured</code></td>
                                                                <td style="padding:6px; font-size: 85%">No effect, the default value of this setting</td>
                                                            </tr>
                                                            <tr>
                                                                <td style="padding:6px; font-size: 85%"><code>Disabled</code></td>
                                                                <td style="padding:6px; font-size: 85%"><code>Restricted</code></td>
                                                            </tr>
                                                            <tr>
                                                                <td style="padding:6px; font-size: 85%"><code>Enabled</code> – Allow only signed scripts</td>
                                                                <td style="padding:6px; font-size: 85%"><code>AllSigned</code></td>
                                                            </tr>
                                                            <tr>
                                                                <td style="padding:6px; font-size: 85%"><code>Enabled</code> – Allow local scripts and remote signed scripts</td>
                                                                <td style="padding:6px; font-size: 85%"><code>RemoteSigned</code></td>
                                                            </tr>
                                                            <tr>
                                                                <td style="padding:6px; font-size: 85%"><code>Enabled</code> – Allow all scripts</td>
                                                                <td style="padding:6px; font-size: 85%"><code>Unrestricted</code></td>
                                                            </tr>
                                                        </table>
                                                    </p>
                                                </ol>
                                            </ul>
                                        </td>
                                    </tr>
                                </table>
                            </ul>
                        </p>
                    <p>For more information, please type "<code>Get-ExecutionPolicy -List</code>", "<code>help Set-ExecutionPolicy -Full</code>", "<code>help about_Execution_Policies</code>" or visit <a href="https://technet.microsoft.com/en-us/library/hh849812.aspx">Set-ExecutionPolicy</a> or <a href="http://go.microsoft.com/fwlink/?LinkID=135170">about_Execution_Policies</a>.</p>
                    </li>
                </p>
                <p>
                    <li><code>New-Item -ItemType File -Path C:\Temp\Get-ExifData.ps1</code><br />
                    Creates an empty ps1-file to the <code>C:\Temp</code> directory. The <code>New-Item</code> cmdlet has an inherent <code>-NoClobber</code> mode built into it, so that the procedure will halt, if overwriting (replacing the contents) of an existing file is about to happen. Overwriting a file with the <code>New-Item</code> cmdlet requires using the <code>Force</code>. If the path name and/or the filename includes space characters, please enclose the whole <code>-Path</code> parameter value in quotation marks (single or double):
                        <ol>
                            <br /><code>New-Item -ItemType File -Path "C:\Folder Name\Get-ExifData.ps1"</code>
                        </ol>
                    <br />For more information, please type "<code>help New-Item -Full</code>".</li>
                </p>
            </ol>
        </td>
    </tr>
</table>




### Contributing

<table>
    <tr>
        <th><img class="emoji" title="contributing" alt="contributing" height="28" width="28" align="absmiddle" src="https://assets-cdn.github.com/images/icons/emoji/unicode/1f33f.png"></th>
        <td style="padding:6px"><strong>Bugs:</strong></td>
        <td style="padding:6px">Bugs can be reported by creating a new <a href="https://github.com/auberginehill/get-exif-data/issues">issue</a>.</td>
    </tr>
    <tr>
        <th rowspan="2"></th>
        <td style="padding:6px"><strong>Feature Requests:</strong></td>
        <td style="padding:6px">Feature request can be submitted by creating a new <a href="https://github.com/auberginehill/get-exif-data/issues">issue</a>.</td>
    </tr>
    <tr>
        <td style="padding:6px"><strong>Editing Source Files:</strong></td>
        <td style="padding:6px">New features, fixes and other potential changes can be discussed in further detail by opening a <a href="https://github.com/auberginehill/get-exif-data/pulls">pull request</a>.</td>
    </tr>
</table>




### www

<table>
    <tr>
        <th>:globe_with_meridians:</th>
        <td style="padding:6px"><a href="https://github.com/auberginehill/get-exif-data">Script Homepage</a></td>
    </tr>
    <tr>
        <th rowspan="32"></th>
        <td style="padding:6px">clayman2: <a href="http://powershell.com/cs/media/p/7476.aspx">Disk Space</a> (or one of the <a href="http://web.archive.org/web/20120304222258/http://powershell.com/cs/media/p/7476.aspx">archive.org versions</a>)</td>
    </tr>
    <tr>
        <td style="padding:6px">Franck Richard: <a href="http://franckrichard.blogspot.com/2011/04/2011-scripting-games-advanced-event-8.html">Use PowerShell to Remove Metadata and Resize Images</a></td>
    </tr>
    <tr>
        <td style="padding:6px">lamaar75: <a href="http://powershell.com/cs/forums/t/9685.aspx">Creating a Menu</a> (or one of the <a href="https://web.archive.org/web/20150910111758/http://powershell.com/cs/forums/t/9685.aspx">archive.org versions</a>)</td>
    </tr>
    <tr>
        <td style="padding:6px">Twon of An: <a href="https://community.spiceworks.com/scripts/show/2263-get-the-sha1-sha256-sha384-sha512-md5-or-ripemd160-hash-of-a-file">Get the SHA1,SHA256,SHA384,SHA512,MD5 or RIPEMD160 hash of a file</a></td>
    </tr>
    <tr>
        <td style="padding:6px">Gisli: <a href="http://stackoverflow.com/questions/8711564/unable-to-read-an-open-file-with-binary-reader">Unable to read an open file with binary reader</a></td>
    </tr>
    <tr>
        <td style="padding:6px">Fred: <a href="https://social.technet.microsoft.com/Forums/scriptcenter/en-US/76ae6430-4993-4422-aa97-8f8ec3ca4e87/selectobject-where?forum=winserverpowershell">select-object | where</a></td>
    </tr>
    <tr>
        <td style="padding:6px"><a href="https://stackoverflow.com/questions/27175137/powershellv2-remove-last-x-characters-from-a-string">Powershell v2 - remove last x characters from a string</a></td>
    </tr>
    <tr>
        <td style="padding:6px"><a href="http://nicholasarmstrong.com/2010/02/exif-quick-reference/">EXIF Quick Reference</a></td>
    </tr>
    <tr>
        <td style="padding:6px"><a href="https://msdn.microsoft.com/en-us/library/ms630826(v=vs.85).aspx">Shared Samples</a></td>
    </tr>
    <tr>
        <td style="padding:6px"><a href="https://sno.phy.queensu.ca/~phil/exiftool/TagNames/EXIF.html">EXIF Tags</a></td>
    </tr>
    <tr>
        <td style="padding:6px"><a href="https://stackoverflow.com/questions/7076958/read-exif-and-determine-if-the-flash-has-fired">Read EXIF and determine if the flash has fired</a></td>
    </tr>
    <tr>
        <td style="padding:6px"><a href="https://technet.microsoft.com/en-us/library/ff730939.aspx">Adding a Simple Menu to a Windows PowerShell Script</a></td>
    </tr>
    <tr>
        <td style="padding:6px"><a href="https://technet.microsoft.com/en-us/library/ee692804.aspx">The String's the Thing</a></td>
    </tr>
    <tr>
        <td style="padding:6px"><a href="http://kb.winzip.com/kb/entry/207/">Snap and Share on Windows Server</a></td>
    </tr>
    <tr>
        <td style="padding:6px"><a href="https://msdn.microsoft.com/en-us/library/windows/desktop/ms630506(v=vs.85).aspx">ImageFile object</a></td>
    </tr>
    <tr>
        <td style="padding:6px"><a href="https://blogs.msdn.microsoft.com/powershell/2009/03/30/image-manipulation-in-powershell/">Image Manipulation in PowerShell</a></td>
    </tr>
    <tr>
        <td style="padding:6px"><a href="https://stackoverflow.com/questions/4304821/get-startup-type-of-windows-service-using-powershell">Get startup type of Windows service using PowerShell</a></td>
    </tr>
    <tr>
        <td style="padding:6px"><a href="https://docs.microsoft.com/fi-fi/powershell/module/Microsoft.PowerShell.Management/Get-WmiObject?view=powershell-5.1">Get-WmiObject</a></td>
    </tr>
    <tr>
        <td style="padding:6px"><a href="https://social.microsoft.com/Forums/en-US/4dfe4eec-2b9b-4e6e-a49e-96f5a108c1c8/using-powershell-as-a-photoshop-replacement?forum=Offtopic">Using Powershell as a photoshop replacement</a></td>
    </tr>
    <tr>
        <td style="padding:6px"><a href="https://msdn.microsoft.com/en-us/library/ms630826(VS.85).aspx#SharedSample012">Display Detailed Image Information</a></td>
    </tr>
    <tr>
        <td style="padding:6px"><a href="https://docs.microsoft.com/fi-fi/powershell/module/Microsoft.PowerShell.Utility/Get-FileHash?view=powershell-5.1">Get-FileHash</a></td>
    </tr>
    <tr>
        <td style="padding:6px"><a href="https://stackoverflow.com/questions/21252824/how-do-i-get-powershell-4-cmdlets-such-as-test-netconnection-to-work-on-windows">How do I get PowerShell 4 cmdlets such as Test-NetConnection to work on Windows 7?</a></td>
    </tr>
    <tr>
        <td style="padding:6px"><a href="https://msdn.microsoft.com/en-us/library/system.security.cryptography.sha256cryptoserviceprovider(v=vs.110).aspx">SHA256CryptoServiceProvider Class</a></td>
    </tr>
    <tr>
        <td style="padding:6px"><a href="https://www.experts-exchange.com/questions/25100459/I-need-to-send-the-details-of-a-jpg-file-to-an-array-any-windows-api-to-do-this-or-get-me-started.html">Send the details of a jpg file to an array</a></td>
    </tr>
    <tr>
        <td style="padding:6px"><a href="https://social.technet.microsoft.com/Forums/windowsserver/en-US/16124c53-4c7f-41f2-9a56-7808198e102a/attribute-seems-to-give-byte-array-how-to-convert-to-string?forum=winserverpowershell">Attribute seems to give byte array. How to convert to string?</a></td>
    </tr>
    <tr>
        <td style="padding:6px"><a href="http://compgroups.net/comp.databases.ms-access/handy-routine-for-getting-file-metad/1484921">Handy routine for getting file metadata</a></td>
    </tr>
    <tr>
        <td style="padding:6px"><a href="http://www.exiv2.org/tags.html">Standard Exif Tags</a></td>
    </tr>
    <tr>
        <td style="padding:6px"><a href="https://sno.phy.queensu.ca/~phil/exiftool/TagNames/GPS.html">GPS Tags</a></td>
    </tr>
    <tr>
        <td style="padding:6px"><a href="https://blogs.technet.microsoft.com/heyscriptingguy/2013/09/21/powertip-use-powershell-to-send-beep-to-console/">PowerTip: Use PowerShell to Send Beep to Console</a></td>
    </tr>
    <tr>
        <td style="padding:6px"><a href="https://stackoverflow.com/questions/21048650/how-can-i-append-files-using-export-csv-for-powershell-2">How can I append files using export-csv for PowerShell 2</a></td>
    </tr>
    <tr>
        <td style="padding:6px"><a href="https://blogs.technet.microsoft.com/heyscriptingguy/2011/11/02/remove-unwanted-quotation-marks-from-csv-files-by-using-powershell/">Remove Unwanted Quotation Marks from CSV Files by Using PowerShell</a></td>
    </tr>
</table>




### Related scripts

 <table>
    <tr>
        <th><img class="emoji" title="www" alt="www" height="28" width="28" align="absmiddle" src="https://assets-cdn.github.com/images/icons/emoji/unicode/0023-20e3.png"></th>
        <td style="padding:6px"><a href="https://gist.github.com/auberginehill/aa812bfa79fa19fbd880b97bdc22e2c1">Disable-Defrag</a></td>
    </tr>
    <tr>
        <th rowspan="28"></th>
        <td style="padding:6px"><a href="https://github.com/auberginehill/emoji-table">Emoji Table</a></td>
    </tr>
    <tr>
        <td style="padding:6px"><a href="https://github.com/auberginehill/firefox-customization-files">Firefox Customization Files</a></td>
    </tr>
    <tr>
        <td style="padding:6px"><a href="https://github.com/auberginehill/get-ascii-table">Get-AsciiTable</a></td>
    </tr>
    <tr>
        <td style="padding:6px"><a href="https://github.com/auberginehill/get-battery-info">Get-BatteryInfo</a></td>
    </tr>
    <tr>
        <td style="padding:6px"><a href="https://github.com/auberginehill/get-bing-background-images">Get-BingBackgroundImages</a></td>
    </tr>
    <tr>
        <td style="padding:6px"><a href="https://github.com/auberginehill/get-computer-info">Get-ComputerInfo</a></td>
    </tr>
    <tr>
        <td style="padding:6px"><a href="https://github.com/auberginehill/get-culture-tables">Get-CultureTables</a></td>
    </tr>
    <tr>
        <td style="padding:6px"><a href="https://github.com/auberginehill/get-directory-size">Get-DirectorySize</a></td>
    </tr>
    <tr>
        <td style="padding:6px"><a href="https://github.com/auberginehill/get-hash-value">Get-HashValue</a></td>
    </tr>
    <tr>
        <td style="padding:6px"><a href="https://github.com/auberginehill/get-installed-programs">Get-InstalledPrograms</a></td>
    </tr>
    <tr>
        <td style="padding:6px"><a href="https://github.com/auberginehill/get-installed-windows-updates">Get-InstalledWindowsUpdates</a></td>
    </tr>
    <tr>
        <td style="padding:6px"><a href="https://github.com/auberginehill/get-powershell-aliases-table">Get-PowerShellAliasesTable</a></td>
    </tr>
    <tr>
        <td style="padding:6px"><a href="https://gist.github.com/auberginehill/9c2f26146a0c9d3d1f30ef0395b6e6f5">Get-PowerShellSpecialFolders</a></td>
    </tr>
    <tr>
        <td style="padding:6px"><a href="https://github.com/auberginehill/get-ram-info">Get-RAMInfo</a></td>
    </tr>
    <tr>
        <td style="padding:6px"><a href="https://gist.github.com/auberginehill/eb07d0c781c09ea868123bf519374ee8">Get-TimeDifference</a></td>
    </tr>
    <tr>
        <td style="padding:6px"><a href="https://github.com/auberginehill/get-time-zone-table">Get-TimeZoneTable</a></td>
    </tr>
    <tr>
        <td style="padding:6px"><a href="https://github.com/auberginehill/get-unused-drive-letters">Get-UnusedDriveLetters</a></td>
    </tr>
    <tr>
        <td style="padding:6px"><a href="https://github.com/auberginehill/get-windows-10-lock-screen-wallpapers">Get-Windows10LockScreenWallpapers</a></td>
    </tr>
    <tr>
        <td style="padding:6px"><a href="https://github.com/auberginehill/java-update">Java-Update</a></td>
    </tr>
    <tr>
        <td style="padding:6px"><a href="https://github.com/auberginehill/remove-duplicate-files">Remove-DuplicateFiles</a></td>
    </tr>
    <tr>
        <td style="padding:6px"><a href="https://github.com/auberginehill/remove-empty-folders">Remove-EmptyFolders</a></td>
    </tr>
    <tr>
        <td style="padding:6px"><a href="https://gist.github.com/auberginehill/13bb9f56dc0882bf5e85a8f88ccd4610">Remove-EmptyFoldersLite</a></td>
    </tr>
    <tr>
        <td style="padding:6px"><a href="https://gist.github.com/auberginehill/176774de38ebb3234b633c5fbc6f9e41">Rename-Files</a></td>
    </tr>
    <tr>
        <td style="padding:6px"><a href="https://github.com/auberginehill/rock-paper-scissors">Rock-Paper-Scissors</a></td>
    </tr>
    <tr>
        <td style="padding:6px"><a href="https://github.com/auberginehill/toss-a-coin">Toss-a-Coin</a></td>
    </tr>
    <tr>
        <td style="padding:6px"><a href="https://github.com/auberginehill/unzip-silently">Unzip-Silently</a></td>
    </tr>
    <tr>
        <td style="padding:6px"><a href="https://github.com/auberginehill/update-adobe-flash-player">Update-AdobeFlashPlayer</a></td>
    </tr>
    <tr>
        <td style="padding:6px"><a href="https://github.com/auberginehill/update-mozilla-firefox">Update-MozillaFirefox</a></td>
    </tr>
</table>
