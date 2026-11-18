# Nine Distribution: NON-DESTRUCTIVE SharePoint Master Migration
# Requires: Install-Module Microsoft.Graph.Authentication -Scope CurrentUser
# First run WITHOUT -Apply (dry run). Then run WITH -Apply to create/copy.
param([switch]$Apply)
$ErrorActionPreference='Stop'
Import-Module Microsoft.Graph.Authentication
Connect-MgGraph -Scopes 'Sites.Manage.All','Sites.ReadWrite.All' -NoWelcome
$site='ninedistribution.sharepoint.com,13d8bc75-67a0-4dbf-b579-8b26c2dd2e9f,dedddf0b-d21b-4660-84ad-259b6ffb0ab4'
$base='https://graph.microsoft.com/v1.0/sites/'+$site
$sources=@(@{kind='Computer';id='ac67c228-aaa2-44f2-8f8e-72d61dd020f0'},@{kind='Accessories';id='7a610ea3-7308-4870-a273-9770713f3594'})
$folder='master-migration-'+(Get-Date -Format yyyyMMdd-HHmmss)
New-Item -Path $folder -ItemType Directory -Force | Out-Null
function CallGraph($url,$verb='GET',$body=$null){
 if($null -eq $body){return Invoke-MgGraphRequest -Uri $url -Method $verb}
 return Invoke-MgGraphRequest -Uri $url -Method $verb -Body ($body|ConvertTo-Json -Depth 20) -ContentType 'application/json'
}
function GetAll($url){
 $result=@()
 while($url){
  $page=CallGraph $url
  $result+=@($page.value)
  $url=$page['@odata.nextLink']
 }
 return ,$result
}
$rows=@()
foreach($source in $sources){
 $uri=$base+'/lists/'+$source.id+'/items?%24top=100&%24expand=fields'
 $items=@(GetAll $uri)
 Write-Host ($source.kind+': '+$items.Count)
 foreach($it in $items){$rows+=@{kind=$source.kind;listId=$source.id;item=$it}}
}
$rows|ConvertTo-Json -Depth 30|Set-Content ($folder+'/original-data.json') -Encoding utf8
$mapped=@()
foreach($entry in $rows){
 $f=$entry.item.fields
 $isComputer=$entry.kind -eq 'Computer'
 $holder=if($isComputer){[string]$f._x0e4a_UserName}else{[string]$f.UserName0}
 $org=if($isComputer){[string]$f.Company}else{[string]$f.Owner}
 $comp=if($org -match '(?i)after|a9'){'After Nine Co.,ltd'}else{'Nine distribution co.,ltd'}
 $hostId=[string]$f.HostName
 $title=if($isComputer){[string]$f.Title}elseif($hostId.Trim()){$hostId}else{'LEGACY-ACC-'+$entry.item.id}
 $serial=if($isComputer){[string]$f.Assetorserialnumber}else{[string]$f.AssetorserialNumber}
 $mapped+=@{
  Title=$title;AssetID=$title;HostName=$hostId;CompanyName=$comp;
  Category=[string]$f.ItemType;Brand=[string]$f.Vendor;Model=[string]$f.Model;
  SerialNumber=$serial;AssetStatus=[string]$f.Status;UserName=$holder.Replace([char]9,' ').Trim();
  OriginalTitle=[string]$f.Title;SourceList=$entry.kind;SourceListId=$entry.listId;
  SourceItemId=[string]$entry.item.id;ReceivedDateText=[string]$f.ReceivedDate;
  ReturnedDateText=[string]$f.ReturnedDate;WarrantyText=[string]$f.Warranty;
  OldUser=[string]$f.OldUser
 }
}
$mapped|ConvertTo-Json -Depth 20|Set-Content ($folder+'/mapped-data.json') -Encoding utf8
$attachmentCount=@($rows|Where-Object {$_.item.fields.Attachments -eq $true}).Count
$collisionCount=@($mapped|Group-Object AssetID|Where-Object Count -gt 1).Count
$ownerCount=@($mapped|Where-Object {$_.UserName}).Count
Write-Host ('Total rows: '+$mapped.Count+'; with users: '+$ownerCount+'; attachments: '+$attachmentCount+'; duplicate AssetID groups: '+$collisionCount)
Write-Host ('Private backup folder: '+$folder)
if(-not $Apply){Write-Host 'DRY RUN — no SharePoint changes. To copy, rerun with -Apply.';return}
$masterName='9D Master Asset Inventory'
$lists=@(GetAll ($base+'/lists?%24top=100'))
$master=$lists|Where-Object {$_.displayName -eq $masterName}|Select-Object -First 1
if(-not $master){
 $master=CallGraph ($base+'/lists') 'POST' @{displayName=$masterName;list=@{template='genericList'}}
}
$id=$master.id
$columns=@(GetAll ($base+'/lists/'+$id+'/columns?%24top=100'))
$names=@('AssetID','HostName','CompanyName','Category','Brand','Model','SerialNumber','AssetStatus','UserName','OriginalTitle','SourceList','SourceListId','SourceItemId','ReceivedDateText','ReturnedDateText','WarrantyText','OldUser')
foreach($name in $names){
 if($columns.name -notcontains $name){[void](CallGraph ($base+'/lists/'+$id+'/columns') 'POST' @{name=$name;text=@{}})}
}
$existing=@(GetAll ($base+'/lists/'+$id+'/items?%24top=100&%24expand=fields'))
$keys=@{};foreach($x in $existing){$keys[([string]$x.fields.SourceListId+'|'+[string]$x.fields.SourceItemId)]=1}
$created=0;$failed=@()
foreach($r in $mapped){
 $key=$r.SourceListId+'|'+$r.SourceItemId
 if($keys.ContainsKey($key)){continue}
 try{[void](CallGraph ($base+'/lists/'+$id+'/items') 'POST' @{fields=$r});$created++}
 catch{$failed+=@{key=$key;error=$_.Exception.Message}}
}
$final=@(GetAll ($base+'/lists/'+$id+'/items?%24top=100&%24expand=fields'))
$index=@{};foreach($x in $final){$index[([string]$x.fields.SourceListId+'|'+[string]$x.fields.SourceItemId)]=$x}
$missing=@();$mismatched=@()
foreach($r in $mapped){
 $key=$r.SourceListId+'|'+$r.SourceItemId
 if(-not $index.ContainsKey($key)){$missing+=$key;continue}
 foreach($field in @('AssetID','CompanyName','SerialNumber','AssetStatus','UserName','Model','HostName')){
  if([string]$r[$field] -cne [string]$index[$key].fields[$field]){$mismatched+=($key+':'+$field)}
 }
}
$report=@{
 SourceRows=$mapped.Count;CopiedNow=$created;TargetRows=$final.Count;MissingSourceKeys=$missing.Count;
 FieldMismatches=$mismatched.Count;AssetIdCollisionGroups=$collisionCount;AttachmentsToReview=$attachmentCount;
 ReadyForCutover=($missing.Count -eq 0 -and $mismatched.Count -eq 0 -and $collisionCount -eq 0 -and $attachmentCount -eq 0)
}
$report|ConvertTo-Json|Set-Content ($folder+'/migration-report.json') -Encoding utf8
$failed|ConvertTo-Json|Set-Content ($folder+'/failed-items.json') -Encoding utf8
$missing|ConvertTo-Json|Set-Content ($folder+'/missing-items.json') -Encoding utf8
$mismatched|ConvertTo-Json|Set-Content ($folder+'/mismatches.json') -Encoding utf8
$report|Format-List
Write-Warning 'Source lists unchanged. Attachment files, version history, item permissions are not copied by this script. DO NOT switch production until they are reviewed.'
