<#
Usage:
.\Get-CalendarSizes.ps1 -filePath c:\temp

Get-CalendarSizes uses Get-MailboxFolderStatistics to collect Identity info and the size of the main user calendar for all mailboxes in an O365 Exchange Online tenant

    THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
    IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
    FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
    AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
    LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
    OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
    SOFTWARE

#>

param(
[parameter(Mandatory=$true)]
[ValidateNotNullOrEmpty()]
[string]$filePath
)

[array]$Mbx = Get-ExoMailbox -RecipientTypeDetails UserMailbox -ResultSize Unlimited | select DisplayName, UserPrincipalName

$OutputReport = [System.Collections.Generic.List[Object]]::new()

foreach ($M in $Mbx){
$stats = Get-EXOMailboxFolderStatistics -UserPrincipalName $m.UserPrincipalName | ?{$_.FolderPath -eq "/Calendar"}

  $ReportLine = [PSCustomObject]@{
    DisplayName       = $M.DisplayName
    UPN               = $M.UserPrincipalName
    FolderSize              = $stats.FolderSize
    FolderAndSubfolderSize = $stats.FolderAndSubfolderSize
    Created                = $stats.CreationTime
   } 
   $OutputReport.Add($ReportLine)
 }

 $OutputReport | Export-Csv $filePath\EXOCalendarStatistics_$((Get-Date).ToString('MM-dd-yyyy')).csv -NoTypeInformation