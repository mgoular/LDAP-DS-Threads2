'################################################################################
'################################################################################
'## Script description                                                         ##
'## Function : how DS Threads in use , LDAP bind time , LDAP sessions for DC   ##                                                                          ##
'## Name      : DS-LdapSessions.vbs                                            ##
'## Version   : 0.3                                                            ##
'## last updated : 2018                                                        ##
'## Language  : VBScript                                                       ##
'## License   : MIT                                                            ##
'## Owner     : M.G                                                            ##
'## Authors   : M.G                                                            ##
'################################################################################
'################################################################################

strComputer = "COMPNAME"
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colDatabases = objWMIService.ExecQuery _
    ("Select * from Win32_PerfFormattedData_NTDS_NTDS")

For Each objADDatabase in colDatabases
    Wscript.Echo "DS threads in use: " & objADDatabase.DSThreadsInUse
    Wscript.Echo "LDAP bind time: " & objADDatabase.LDAPBindTime
    Wscript.Echo "LDAP client sessions: " & objADDatabase.LDAPClientSessions
Next
