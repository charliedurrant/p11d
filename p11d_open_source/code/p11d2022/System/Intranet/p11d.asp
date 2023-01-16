<%@LANGUAGE=JSCRIPT %>

<!-- change the m_authentication to one of the constant values in the p11d.inc file starting with L_AUTHENTICATION_XXX -->
<!-- #INCLUDE FILE="p11d.inc" -->

<%
 var ret;
 ret = main();

 if ( ( ret == null ) || (! ret) )
 {
   switch( m_authentication )
   {
     case L_AUTHENTICATION_XML:
%>
      <!-- START XML Authenticate page, lookup password in xmlfile -->
      <html>
        </head>
          <title>P11D information</title>
          <style>
          	BODY
          	{
          		font-family:Verdana;
          		font-size:11pt;
          	}
            TABLE.tablebanner
            {
              cell-padding=0;
              border-style=none;
              padding:10px;
              border-width=0px;
              font-weight=bold;
              color=#ffffff;
              border-top=10px;
            }

            TD.tdbanner
            {
            	padding:10px;
            	background-color:{BANNER_BACK_COLOR};
            	color:{BANNER_FORE_COLOR};
            }

            font.fontheader
            {
              font-weight=bold;
              font-size=20pt;
            }
          </style>
        </head>
        <body style="border-width:0">
          <TABLE cellpadding=0 cellspacing=0 border=0 WIDTH="100%" class="tablebanner">
            <TR>
              <TD class="tdbanner">
                {BANNER_TITLE_HTML}
              </TD>
            </TR>
          </TABLE>
          <br>
          <br>
          <center>
            <font class="fontheader">P11D information</font>
          </center>
          <TABLE WIDTH="100%">
            <TR>
              <TD>
                <img  src="images\logo.jpg"></img>
              </TD>
              <TD align="CENTRE">
                <form name = "login" method="post" ACTION = "p11d.asp" >

                  <table border="0" >
                    <tr>
                      <td>
                        <% logininfo() %>
                      </td>
                    </tr>
                    <tr>
                      <td>
                        Enter username
                      </td>
                      <td>
                        <input type = "text" name = "username" value = "<% si ("username") %>"  size = "30">
                      </td>
                    </tr>
                      <td>
                        Enter password
                      </td>
                      <td>
                        <input type = "password" name = "password" value = "<% si ("password") %>"  size = "30">
                      </td>
                    </tr>
                    <tr>
                     <td COLSPAN="2">
                        <input type = "submit" value = "Login">
                     </td>
                    </tr>

                  </table>
                </form>
              </TD>

            </TR>

          </TABLE>

					<SPAN>
							{USER_INFORMATION_HTML}
          </SPAN>
          <SPAN>
            </BR></BR>
            <I>Note: to print the P11D after logging-in please use the Print button at the top of the page (not File - Print from the menu).  For best results set the print margins using the File - Page Setup menu option (set the top and bottom margins to 0mm and the left and right margins to a maximum of 9mm).  Remove header and footer entries.</I>
          </SPAN>
        </body>
      </html>
      <!-- END XML Authenticate page -->
<%
      break;
     case L_AUTHENTICATION_WINDOWS:

%>
      <!-- START Windows Authentication page, no username / password box required -->
      <html>
         <head>
           <title>No P11D</title>
         </head>
         <body>
           No P11D is available for this period.
         </body>
      </html>
      <!-- END Windows Authentication page -->
<%
      break;
     case L_AUTHENTICATION_OTHER:

%>
      <!-- START Other Authentication page, no username / password box required -->
      <html>
         <head>

         </head>
         <body>
           Place other authentication html here
         </body>
      </html>
      <!-- END Other Authentication page -->

<%
   }
 }
%>



