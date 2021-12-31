get-appxpackage -allusers *weather* | remove-appxpackage
get-appxpackage -allusers *mixedreality* | remove-appxpackage
get-appxpackage -allusers *getstarted* | remove-appxpackage
get-appxpackage -allusers *3dViewer* | remove-appxpackage
get-appxpackage -allusers *FeedbackHub* | remove-appxpackage
get-appxpackage -allusers *gethelp* | remove-appxpackage
get-appxpackage -allusers *zune* | remove-appxpackage

$appsRemove = @('weather', 'mixedreality', 'getstarted', '3dViewer','FeedbackHub','GetHelp','Zune')
foreach ($remove in $appsRemove) {
    get-appxpackage -allusers *$remove* | remove-appxpackage
}