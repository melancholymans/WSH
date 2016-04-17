var fso,ts;

fso = WScript.CreateObject("Scripting.FileSystemObject");
ts = fso.OpenTextFile("C:\\myproject\\learningWSH\\sample.txt",8,-2);
WScript.Echo(ts.WriteLine("\nRasbeeryPi"));
ts.Close();


