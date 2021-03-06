username=inputbox("Enter username:")
IF username = "" THEN wscript.quit

ldapPath = FindUser(username)

IF ldapPath = "Not Found" THEN
	wscript.echo "User not found!"
ELSE
	SET objUser = GETOBJECT(ldapPath)
	IF isAccountLocked(objUser) THEN
		objuser.put "lockoutTime", 0
		objUser.setinfo
		wscript.echo "Account Unlocked"
	ELSE
		wscript.echo "This account is not locked out"
	END IF
END IF


FUNCTION FindUser(BYVAL UserName) 
	ON ERROR RESUME NEXT

	SET objRoot = GETOBJECT("LDAP://RootDSE")
	domainName = objRoot.GET("defaultNamingContext")
	SET cn = CREATEOBJECT("ADODB.Connection")
	SET cmd = CREATEOBJECT("ADODB.Command")
	SET rs = CREATEOBJECT("ADODB.Recordset")

	cn.open "Provider=ADsDSOObject;"
	
	cmd.activeconnection=cn
	cmd.commandtext="SELECT ADsPath FROM 'LDAP://" & domainName & _
			"' WHERE sAMAccountName = '" & UserName & "'"
	
	SET rs = cmd.EXECUTE

	IF err<>0 THEN
		wscript.echo "Error connecting to Active Directory Database:" & err.description
		wscript.quit
	ELSE
		IF NOT rs.BOF AND NOT rs.EOF THEN
     			rs.MoveFirst
     			FindUser = rs(0)
		ELSE
			FindUser = "Not Found"
		END IF
	END IF
	cn.close
END FUNCTION

FUNCTION IsAccountLocked(BYVAL objUser)
    	ON ERROR RESUME NEXT
	SET objLockout = objUser.GET("lockouttime")

	IF err.number = -2147463155 THEN
		isAccountLocked = FALSE
		EXIT FUNCTION
	END IF
	ON ERROR GOTO 0
	
	IF objLockout.lowpart = 0 AND objLockout.highpart = 0 THEN
		isAccountLocked = FALSE
	ELSE
		isAccountLocked = TRUE
	END IF

END FUNCTION