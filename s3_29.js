var fso,ts;

fso = WScript.CreateObject("Scripting.FileSystemObject");
ts = fso.OpenTextFile("C:\\myproject\\learningWSH\\sample.txt",1);
WScript.Echo(ts.ReadLine());
ts.Close();

