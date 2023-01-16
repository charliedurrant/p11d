<%

var L_AUTHENTICATION_XML = 0;
var L_AUTHENTICATION_WINDOWS = 1;
var L_AUTHENTICATION_OTHER = 2;

//change the authentication here;

var m_authentication = {AUTHENTICATION_TYPE};

var L_MORE_THAN_1_USER = 1;
var L_HTM_FILE_DOES_NOT_EXIST=2;
var m_xml = null;
var m_case_sensitive = {CASE_SENSITIVE};
function si(s)
{ Response.Write ( Session ( s )); }

function formtosession()
{
  try
  {
    Session( "username" ) = Request.Form( "username" )();
    Session( "password" ) = Request.Form( "password" )();
  }
  catch( e )
  { ; }
}

function getuser(username)
{
  var dom;
  var selectstring;
  var users; 
    
  try
  {     

    dom = xml();        
    
    if ( ! m_case_sensitive )
    {
      username = username.toUpperCase();
    }
    
    selectstring = "EMPLOYERS/EMPLOYER/EMPLOYEE[@username='" + username + "']";        
    users = dom.selectNodes(selectstring);        
    
    if ( users == null )
    {       
      return null; 
    }
    if ( users.length > 1 )
    { throw new Error ( L_MORE_THAN_1_USER, "More than one available entry for username: " + username); }
    
    if ( users.length == 0 )
    { 
      return null;
    }
    else
    {             
      return users.nextNode();
    }
        
  }
  catch ( e )
  { throw (e); }
} 

//called directly form p11d.asp
function logininfo()
{
  if ( Session( "loginattempts") == 0 )
  { Response.Write ( "Please enter your details below:" ) }
  else
  { Response.Write ( "Login failed, please try again." ) }

  Session ( "loginattempts" ) = Session ( "loginattempts" ) + 1;
}

function windows_authentication_user(remove_domain)
{
  var s,arr;
  s = new String(Request.ServerVariables("LOGON_USER"));
  if ( s.length == 0 )
  { return s; }
  
  if ( ! remove_domain )
  { return s; }
  
  arr = s.split("\\");  
  return arr[arr.length - 1];   
}

function validatelogin( username, password  )
{
  var user;
    
  try
  {                  
    switch ( m_authentication )
    {
      case L_AUTHENTICATION_WINDOWS:
        user = getuser( windows_authentication_user(true) );                                        
        return user;        
      case L_AUTHENTICATION_XML:
      {
        if ( ( username == null ) || ( username.length == 0 ) )
        { 
          Response.Write ("You have to enter a username");
          return;  
        }                
        user = getuser( username );
        if ( user == null )
        { 
          Response.Write ("User not found");
          return; 
        }     
        
        if ( ! m_case_sensitive )
        {        
          if ( (user.getAttribute ( "password" ).toUpperCase()) == (password.toUpperCase()) )
          { return user; }
        }
        else
        {
          if ( user.getAttribute ( "password" ) == password )
          { return user; }
        }
        
        Response.Write ("Your password is invalid");
        return null; 
        
      }  
      case L_AUTHENTICATION_OTHER:
        /* place custom validation in here, also change HTML on p11d.asp accordingly */
        break;
    }      
  }
  catch ( e )
  { throw (e); }
}

function fileexists( filename )
{
  var fso; 
  var s;
  var sdirpath;
  var sfile;
  
  try
  {
    fso = new ActiveXObject("Scripting.FileSystemObject");
    s = getfileroot (filename);   
    sdirpath = fullpath( s );    
    sfile = sdirpath  + "\\" + filename;        
    if (!fso.FileExists( sfile ) )
    { throw new Error (L_HTM_FILE_DOES_NOT_EXIST, "The p11d html file does not exist for this username, although the user is present in the XML file. ");}
    return sfile;    
    
  }
  catch ( e )
  { throw (e); }
}


function getfileroot ( filename )
{
  var sfile;
  var dom;
  var svirtualroot;
  
  try
  { 
    dom = xml();
    svirtualroot = dom.selectSingleNode("EMPLOYERS/EMPLOYER[EMPLOYEE/@filename ='" + filename + "']").getAttribute ( "virtualroot" );
    sfile = Server.MapPath( svirtualroot );
    
    return sfile; 
    
  }
  catch ( e )
  { throw (e); }  
} 


function fullpath(virtualroot)
{
  var length;
  var newroot;
  
  try
  {
    if ( virtualroot == null )
    { throw new Error( "virtualroot is null" ); }
    
    length = virtualroot.length;
    if (virtualroot.substr ( length - 1, 1 ) == "\\")
    { newroot = virtualroot.substr ( 0, length - 1 ); }
    else
    { newroot = virtualroot; }
    
    return newroot;
  }
  
  catch ( e )
  { throw ( e );}
}

function filetext( spathandfile )
{
  var fso;
  var ts;
  
  try
  {    
    fso = new ActiveXObject( "Scripting.FileSystemObject" );
    ts = fso.OpenTextFile( spathandfile, 1 );    
    return ts.ReadAll();
  }
  catch( e )
  { throw ( e ); }
}

function xml()
{ return Application("dom"); }

function main()
{
  var err;
  var user;
  
  Response.Expires = -1;  
 
  err = xml().parseError; 
  if (err.errorCode != 0) 
  { 
    Response.Write( err.description ); 
    return;
  }  
  if ( ( Session("loginattempts") != 0 ) || (m_authentication == L_AUTHENTICATION_WINDOWS) )
  {
    try
    {            
      formtosession();        
      
      user = validatelogin( Session( "username" ),Session( "password" ) );                
      
      if ( user != null )
      {         
        
        Response.Write( filetext( fileexists( user.getAttribute("filename") )  ) );      
        return true;
      }            
      
    }
    catch ( e )
    { Response.Write (e.description);}
  }    
}



function xml()
{
  if ( m_xml == null )
  {
    m_xml = new ActiveXObject("Msxml.FreeThreadedDOMDocument");
    m_xml.load(Server.MapPath("{XML_FILE_NAME_NO_PATH}"));    
  }  
  return m_xml;
}

%>