<#
.SYNOPSIS
	Installs Microsoft Teams
.DESCRIPTION
	This PowerShell script installs Microsoft Teams from the Microsoft Store.
#>

echo "
                                                                                                    
                                                                                                    
                                                                                                    
                                      .:^~~!!!77777777!!!~~^:..                                     
                                .^~7777!!!~~~~~^^^^^^^~~~~~!!7777~^.                                
                            :~7?77!~~~~~~~~~~~~~~~~~~~~~~~~~~^^^~~!777~.                            
                        .^7?7!~~~~~~!!!!!!!!!!!!!!!~~~~~!!!!!!!~~~~^~~!7?!:                         
                      :7?7!~~~~!!!!!!7777!!!~~^^^^:::^~~~~^^~!!!!!!!!~~~~~7?!:                      
                   .~??!~~!!!!!77777!!~~^^:::::::::::::::::::::^^^^^~~~!!~~~!??~                    
                 .!J?!~!!!!777777!~^^::::::::::::::::::::::::::::::::::^^!!!!~~7J!.                 
                !J?!!!!!777777~^^::::::::::::::::::::::::::::::::::::^^~~^^!!!~~~7J~                
              ^J?!!!!7777???7::::::::::::::::::::::::::::::::::::::::::::~777!!!!~!?J:              
             7J7!!!777~^^~!7~:::::::::::::::::::::::::::::::::::::::::::^?777777!!!~!J!             
           .JJ!!777777!~^:::::::::::::::::::::::::::::::::::::::::::::::^^^^^^~!77!!~!??.           
          .J?!!7777!~^::::::::::::::::::::::::::::::::::::::::::::::::::::::::^!777!!!!?J.          
         .J?!7777?77~::::::::::::::...:::::::::::::::::::::....::^~::::::::::::~77777!!!?J.         
         JJ7777?????^:::::::::::.5B5?!~^^::::::::::::::^^~!7J5PB#&&7.::::::::::^7??7777!!??         
        7J7777??????!:::::::::.^P&&&&&&&##BBBGGGGGBBB###&&&&&&&&&&&&5^.::::::::!????7777!!J!        
       :Y?77?????JJJ?^:::::::.?&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&#7:::::::~?J?????777!7Y.       
       7J77?????JJJJJ7^::::::Y&&&&&&&&&&&&&&&&##&&&&&##&&&&&&&&&&&&&&&&!.::::^7JJJ?????777!J7       
      .J?77????JJJJJJJ!:::::~&&&&&&&&&&&#G5J7~!75&&&#?7~!?YPB#&&&&&&&&&?.:::^7JJJJJ????7777?J       
      :Y?7????JJJJ5BBBP!:::.!&&&&&&#GY7~~~~!7?YJB&&&#5JJ7!!~~~!?PB&&&&&?::::75GBBPJJ????7777Y:      
      ~Y?????JJJJJG&@&BB7::.!&&&&GJ?J5G##&&&&&&&&&&&&&&&&&&&&#BGYJ?5#&&7:::?BB#@@BJJJ????777Y^      
      ~Y?????JJJJJG#####B!::^&&&PP#&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&B5#&!::?B###&#BJJJ?????77Y~      
      ^Y????JJJJJJGBBBGPPP:.:!!!!77777?????JJJJYY5PP55YYY55555555555555J?JGPPGBBBBJJJJ????77Y^      
      .Y????JJJJJJYYYYYJJP~...?GGGGP5555PPPP55?:^:^^^~7BGGGGGGGGPPGBBB5!YJYJJYYYYYJJJJ????7?Y:      
       JJ???JJJJJJJJJJYYY5Y!^.5&&&&@&&B##GGG&&5..^!!:.!@&&G#BGGB&&@&&&5:!??JYYJJJJJJJJ????7?J       
       !J???JJJJJJJYYYYYYYY5Y.^&&&&@@@#Y!^~B@@!.?&&#B.:&@@BY~^!&@@@&&&7.7YYYYYYJJJJJJJ?????J!       
       .JJ??JJJJJJJYYYYYYYYY5~:7#&&&@@@@@@@@@J.~&&&#&Y.~#@@@@@@@@@&&#J.~YYYYYYYJJJJJJJ?????Y.       
        ~Y??JJJJJJJYYYYYJY555?J!~JPB######GY!7P&&&&#&&G7~?PB#####BPJ!!JY555YJJYJJJJJJJ????Y!        
         ?J?JJJJJJJJYYYYG#&&&&&&B5J???JJJJYPB&&&&&&&#&&&#PYJJJJ?JJYPB&&#&&&&BYJJJJJJJJ???J?         
          ?JJJJJJJJJYYYB&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&##&&&&&&&&&&&&&&&&&&&&&#YJJJJJJJ??JJ.         
          .?JJJJJJJJYYY#&&&&&&&&&&&&&&&&&&&&&&&&&&&&&###&&&&&&&&&&&&&&&&&&&&&&YJJJJJJ??J?.          
            7YJJJJJJJYJ5#&&&&&#&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&#&&&&&#5JJJJJJ??J7            
             ~YJJJJJJJYJY5PGGPY#&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&#YPGGP5YJJJJJJ?JY~             
              .?YJJJJJJYJJJJJJJG&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&BJJJJJJJJJJJJJJ?:              
                ^JYJJJJJJYYYYYYY#&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&PJYJJJJJJJJJJJ^                
                  ~JYJJJJJJYYYYY5B&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&#PJYJJJJJJJJYJ~                  
                    ^?YYJJJJJJYYJJ5G#&&&&&&&&&&&&&&&&&&&&&&&&&&&#BPYJJJJJJJJJJ?^                    
                      .!JYYJJJJJYYJJY5PB##&&&&&&&&&&&&&&&&&##BG5YJJJJJJJJJYJ!.                      
                         :!?JYYJJJJJJJJJY55PGBB#######BBGPP5YJJJJJJJJJYJ?!:                         
                            .^!?JYYYJJJJJJJJJYYYYY55YYYYJJJJJJJJYYYJ?!^.                            
                                .:~!?JJYYYYYYYYJJJJJJYYYYYYYYJJ?!~:.                                
                                      ..:^~!!777????777!!~^:..                                      
                                                                                                    
                                                                                                    
                                                                                                     
"

$url32 = 'https://go.microsoft.com/fwlink/p/?LinkID=2187217&culture=en-us&country=WW&Lmsrc=groupChatMarketingPageWeb&Cmpid=directDownloadv2Win32'
$url64 = 'https://go.microsoft.com/fwlink/p/?LinkID=2187217&culture=en-us&country=WW&Lmsrc=groupChatMarketingPageWeb&Cmpid=directDownloadv2Win64'

if(!$url){
    if([Environment]::Is64BitOperatingSystem){
        $url = $url64
    }
    else{
        $url = $url32
    }
}

$path = "c:\teams.exe"

try {
	"Installing Microsoft Teams, please wait..."

	Invoke-WebRequest -URI $url -OutFile $path 
	Start-Process $path
	if ($lastExitCode -ne "0") { throw "' install' failed" }

	"Microsoft Teams installed successfully."
	exit 0 # success
} catch {
	"⚠️ Error in line $($_.InvocationInfo.ScriptLineNumber): $($Error[0])"
	exit 1
}
