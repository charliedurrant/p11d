var m_shell = null;
var m_filelist = "filelist.txt";
var m_iscscript = false;

try
{
  var projects = null;
  var project = null;
  var scriptpath = null;

  String.prototype.trim = trim_string;

  m_shell = WScript.CreateObject("WScript.Shell");
  scriptpath = getscriptpath();
  m_iscscript = iscscript();

  if ( !m_iscscript )
  { WScript.Echo ("This process may take some time!!\nMore feedback is given if run from command prompt" ) ; }

  projects = getprojectfiles(scriptpath + "\\filelist.txt",scriptpath + "\\");

  processprojects(projects)
  
  WScript.Echo ("\n*************** COMPLETED COMPILATION******************\n\n") ;

}
catch (e)
{ 
  WScript.Echo ("\n*************** FAILED COMPILATION ******************\n\n" + e) ;
}

function processprojects(projects)
{
  var i;    
  var project;
  var vbpath = "";
  var scompile = "";
  var fs = null;
  var winsyspath = "";
  var ret = 0;
  
  try
  {
    fs = WScript.CreateObject("Scripting.FileSystemObject");
    winsyspath = fs.GetSpecialFolder(1).Path + "\\";

    vbpath = "\"" + getvbpath() + "\"" ;

    for ( i = 0 ; i < projects.length; i++)
    {
      project = projects[i];

      scompile  = vbpath + " " + project.commandswitch + " " + "\"" + project.pathproject + project.vbp + "\"";
      echoex("\n\ncompiling: " + project.vbp);
      ret = m_shell.Run(scompile,8,true);
      
      if ( ret )       
      { throw "Failed to compile " + project.vbp; }
    
      if ( ret )       
      { throw "Failed to register " + winsyspath + project.component; }


    }

  }
  catch ( e )
  {
    throw "processprojects:" + e ;
  } 
  
  return;

}

function echoex(smsg)
{
  if ( m_iscscript ) 
  { WScript.Echo ( smsg ); }
} 

function iscscript()
{
  return (WScript.Fullname.toLowerCase().indexOf("cscript.exe", 0) != -1);
}

function getscriptpath()
{
  var fs = null;

  fs = WScript.CreateObject("Scripting.FileSystemObject"); 
  return  fs.GetParentFolderName(WScript.ScriptFullName).toLowerCase();
  
}


function getprojectfiles(spathandfile,scurrentpath)
{
  var projects = new Array();
  var ts = null;
  var fs = null;
  var sline = new String();
  var sfile = null;
  var o;
  var s;
  var arr;

  fs = WScript.CreateObject("Scripting.FileSystemObject"); 

  ts = fs.OpenTextFile(spathandfile, 1, false);   
  echoex ("\nReading list of project files\n") ;  

  try
  {
    while (! ts.AtEndOfStream )
    {
      sline = ts.ReadLine().trim();
        
      if ( ( sline.length > 0 ) && ( sline.substring(0, 1) != ";" ) )
      {
        echoex(sline);

        arr = sline.split(";");
      
        if ( arr.length != 3 )
        { throw "The line " + sline + " is not in the correct format of eg stat;tcsstat.vbp;tcsstat.ocx"; }
        
        o = new Object();

        //o.pathproject = scurrentpath + arr[0] + "\\";
        o.pathproject = scurrentpath;

        if ( ! fs.FolderExists(o.pathproject) ) 
        { throw "The project folder " + sfile + " does not exist."; }
        
        //o.pathrelease = o.pathproject + "release\\";
        o.pathrelease = o.pathproject+"\\";
        
        if ( ! fs.FolderExists(o.pathrelease) ) 
        { throw "The release directory " + o.pathrelease + " does not exist."; }

      
        o.vbp =  arr[1];

        if ( ! fs.FileExists(o.pathproject + o.vbp) ) 
        { throw "The project file " + o.pathproject + o.vbp + " does not exist."; }
              
        o.component = arr[2].toLowerCase();
        
        s = o.component.substring(o.component.lastIndexOf ("."),o.component.length );
      
        switch (s)
        {
          case ".ocx":
          case ".dll":
            o.commandswitch = "/l";
            break;
          case ".exe":
            o.commandswitch = "/m";
            break;
          default:
            throw "Invalid component extension: " + s ;    
        }
  
        projects[projects.length] = o;
                    
      }
    }
    ts.close();
    return projects;
  }
  catch( e ) 
  { throw "getprojectfiles: " + e;}

}

function getvbpath()
{
  var s;
  var i;

  try
  {
// xRegGetKeyStr( buffer, HKEY_LOCAL_MACHINE,"Software\\Microsoft\\VisualStudio\\6.0\\Setup\\Microsoft Visual Basic", "ProductDir", 0); 

    s = m_shell.RegRead("HKCR\\VisualBasic.Form\\DefaultIcon\\"); 

    i = s.lastIndexOf(",")
    return  s.substring(0, i);
  }
  catch (e)
  { throw "getvbpath:" + e; } 
} 

function getlineparams(s)
{
  var p0,p1;
  var arr;

  arr = s.split(";");
  
}


function trim_string() 
{
  var ichar, icount;
  var strValue = this;
  ichar = strValue.length - 1;
  icount = -1;
  while (strValue.charAt(ichar)==' ' && ichar > icount)
    --ichar;
  if (ichar!=(strValue.length-1))
    strValue = strValue.slice(0,ichar+1);
  ichar = 0;
  icount = strValue.length - 1;
  while (strValue.charAt(ichar)==' ' && ichar < icount)
    ++ichar;
  if (ichar!=0)
    strValue = strValue.slice(ichar,strValue.length);
  return strValue;
}
