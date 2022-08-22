$svnLocalRoot = "C:\_cp\Documents"
$svnLocalRelativePath = "Documents\Personal"
$svnLocalPath = Join-Path $svnLocalRoot $svnLocalRelativePath
$gitRoot = "C:\_cp\Git"
$gitRepoName = "Money"

#
# Get the authors from the SVN repository
#

Write-Host ( "Looking up authors from SVN repository {0}" -f $svnLocalPath )
cd $svnLocalPath
$logXml = [xml](svn log --xml)

$authors = @{
    "Craig" = "Craig Paton <craigpaton@gmail.com>"
    "craig_000" = "Craig Paton <craigpaton@gmail.com>"
}

$unknownAuthor = $false
$logEntries = $logXml.log.logentry
foreach ( $logEntry in $logEntries ) {
    if ( !$authors.ContainsKey( $logEntry.author ) ) {
        $authors[$logEntry.author] = "Craig Paton <craigpaton@gmail.com>"
        $unknownAuthor = $true
    }    
}

$authorsFile = @()
foreach ( $author in $authors.Keys ) {
    $authorsFile += "{0} = {1}" -f $author, $authors[$author]
}
$authorsFilePath = Join-Path $gitRoot "authors.txt"
Set-Content -Path $authorsFilePath -Value $authorsFile

if ( $unknownAuthor ) {
    $sublimeCommand = '& "C:\Program Files\Sublime Text 3\subl.exe" -wait "{0}"' -f $authorsFilePath
    Invoke-Expression $sublimeCommand
}

#
# Start up a SVN server that the git-svn command can connect to
#

$svnInfo = [xml](svn info $svnLocalRoot --xml)
$svnRemoteRootUrl = $svnInfo.info.entry.repository.root
$svnRemoteRoot = (New-Object -Type System.Uri $svnRemoteRootUrl).LocalPath


Write-Host ( "Starting SVN host for {0}" -f $svnRemoteRoot )
$svnServerScriptBlock = { svnserve -d -R -r $args[0] }.GetNewClosure()
$svnServeJob = Start-Job -ScriptBlock $svnServerScriptBlock -ArgumentList @( $svnRemoteRoot )

#
# Remove any existing repository so we start from scratch
#

Set-Location $gitRoot
if ( Test-Path $gitRepoName ) {
    Remove-Item -Recurse -Path $gitRepoName -Confirm
}

#
# Migrate SVN to Git
#
$gitCloneCommand = 'git svn clone --authors-file="{0}" --prefix="origin/" "svn://localhost/{1}" {2}' -f $authorsFilePath, $svnLocalRelativePath.Replace( "\", "/" ), $gitRepoName
Write-Host $gitCloneCommand
Invoke-Expression $gitCloneCommand

#
# Stop the SVN server
#
$svnServeJob.StopJob()
Receive-Job $svnServeJob

# git branch --delete -r git-svn
# git remote add origin ssh://Craig@storage:60022/volume1/Data/Git/Repos/CAPCON