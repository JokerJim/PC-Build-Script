Set-ExecutionPolicy remotesigned
function downloadbuilder {
Invoke-WebRequest -Uri "https://github.com/jokerjim/PC-Build-Script/archive/master.zip" -OutFile "C:\Pirum\PCBuild.zip"
Expand-Archive C:\Pirum\PCBuild.zip -DestinationPath C:\Pirum\
}
downloadbuilder