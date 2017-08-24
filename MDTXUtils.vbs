
Option Explicit

Dim oMDTXUtils
Set oMDTXUtils = New MDTXTools

Class MDTXTools

    Sub LoadVarsFromXmlFile(sArg)
        Dim oVarList
        Dim oNode
        Dim sVar
        Dim sValue
    	set oVarList = oUtility.CreateXMLDOMObjectSafe(sArg)
		For each oNode in oVarList.selectNodes("//var")
			sVar = oNode.GetAttribute("name")
			sValue = oNode.Text
			oLogging.CreateEntry "Found " & sVar & "=" & sValue, LogTypeInfo
			oEnvironment.Item(sVar)= sValue		
		Next
    End Sub
    
    Sub LogXmlError(oXml)
        oLogging.CreateEntry "Transformation failed:" & oXml.parseError.ErrorCode, LogTypeError
        oLogging.CreateEntry "Reason: " & oXml.parseError.reason, LogTypeError
        oLogging.CreateEntry "Line: " & oXml.parseError.line & " Position: " & oXml.parseError.linepos, LogTypeError
        oLogging.CreateEntry "FilePos: " & oXml.parseError.filepos , LogTypeError
        oLogging.CreateEntry "Source Text: " & oXml.parseError.srcText , LogTypeError           
    End Sub
    
    Function  RunXSD(sXml,sXsd)
        Dim oXml
        Dim oXsd
        Dim oSchemas
        'Dim sNameSpace
        Set oXml = CreateObject("MSXML2.DOMDocument.6.0")
        'oXml.async = False
        Set oXsd = CreateObject("MSXML2.DOMDocument.6.0")
        'oXsd.async= False

        Set oSchemas = CreateObject("MSXML2.XMLSchemaCache.6.0")
        oXml.schemas = oSchemas
        'sNameSpace = oXsd.selectSingleNode("//@targetNamespace").text
        'oSchemas.add sNameSpace, oXsd
        oXsd.load sXsd
        oSchemas.add "", oXsd
        oXml.load sXml
        
        'oXsd.resolveExternals = True
     
        If oXml.parseError.errorCode= 0 Then
        	RunXSD = ""
        Else
            LogXmlError oXml
        	RunXSD = oXml.ParseError.reason & "[" & oXml.ParseError.srcText & "]"
        End If	        
    End Function 
    
    Sub RunXSL(sXml,sXsl,sOut)

		Dim oXml
		Dim oXsl
		Dim oRes
		Dim sRes
        Dim oXml2
		Set oXml =  oUtility.CreateXMLDOMObjectSafe( sXml )
		Set oXsl =  oUtility.CreateXMLDOMObjectSafe( sXsl )
		'Set oRes = CreateObject("MSXML2.DOMDocument")
        Set oRes = oUtility.CreateXMLDOMObject
        Set oXml2 = oUtility.CreateXMLDOMObject
        oLogging.CreateEntry "Applying XSL " & sXsl & " to " & sXml, LogTypeInfo
        oXml.transformNodeToObject oXsl, oXml2
        if oXml.ParseError.ErrorCode=0 Then
            oLogging.CreateEntry "-- Beginning of resulting transform --", LogTypeInfo
                    'for debugging
            sRes = oXml.transformNode(oXsl)	
            oLogging.CreateEntry sRes, LogTypeInfo
            oLogging.CreateEntry "-- End of resulting transform --", LogTypeInfo
            oXml2.Save sOut
        Else
            LogXmlError oXml
        End If

	End Sub
    
       Sub RunXSL2(sXml,sXsl,sOut)

		Dim oXml
		Dim oXsl
		Dim oRes
		Dim sRes
        Dim fRes
        Dim oXml2
		Set oXml =  oUtility.CreateXMLDOMObjectSafe( sXml )
		Set oXsl =  oUtility.CreateXMLDOMObjectSafe( sXsl )
		'Set oRes = CreateObject("MSXML2.DOMDocument")
        Set oRes = oUtility.CreateXMLDOMObject
        Set oXml2 = oUtility.CreateXMLDOMObject
        oLogging.CreateEntry "Applying XSL " & sXsl & " to " & sXml, LogTypeInfo
        'oXml.transformNodeToObject oXsl, oXml2
        if oXml.ParseError.ErrorCode=0 Then
            oLogging.CreateEntry "-- Beginning of resulting transform --", LogTypeInfo
                    'for debugging
            sRes = oXml.transformNode(oXsl)	
            oLogging.CreateEntry sRes, LogTypeInfo
            oLogging.CreateEntry "-- End of resulting transform --", LogTypeInfo
            Set fRes=oFso.CreateTextFile(sOut, True)
            fRes.WriteLine sRes
            fRes.Close
            'sRes.Save sOut
        Else
            LogXmlError oXml
        End If

	End Sub
    
	Function ServiceExists ( ServiceName )
		Dim colServices		
		Dim objService
		Dim objWMIService
		Dim strQuery
		Dim bFound
		
		bFound = False
		
		Set objWMIService = GetObject( "winmgmts:\\.\root\CIMV2" )
		strQuery = "SELECT * FROM Win32_Service WHERE Name='" & ServiceName & "'"
		Set colServices = objWMIService.ExecQuery( strQuery, "WQL", 48 )	
		For Each objService In colServices
			bFound = True
		Next
		ServiceExists = bFound		
	End Function
	
	Sub RestartService( ServiceName )

		Dim colServices
		Dim colSvcCheck
		Dim objService
		Dim objSvcCheck
		Dim objWMIService
		Dim strQuery
		Dim strState

		Set objWMIService = GetObject( "winmgmts:\\.\root\CIMV2" )
		strQuery = "SELECT * FROM Win32_Service WHERE Name='" & ServiceName & "'"
		Set colServices = objWMIService.ExecQuery( strQuery, "WQL", 48 )

		For Each objService In colServices
            oLogging.CreateEntry "Stopping " & ServiceName , LogTypeInfo
			objService.StopService

			strState = ""	
			Do Until strState = "Stopped"
			' Force Updated Info...
				Set colSvcCheck = objWMIService.ExecQuery( strQuery, "WQL", 48 )
				For Each objSvcCheck In colSvcCheck
					strState = objSvcCheck.State
					oLogging.CreateEntry "Waiting Service to stop:" & strState, LogTypeInfo
					WScript.Sleep 1000
				Next
				Set colSvcCheck = Nothing
			Loop
		
			oLogging.CreateEntry "Starting " & ServiceName, LogTypeInfo
			objService.StartService

			strState = ""	
			Do Until strState = "Running"
				Set colSvcCheck = objWMIService.ExecQuery( strQuery, "WQL", 48 )
				For Each objSvcCheck In colSvcCheck
					strState = objSvcCheck.State
					oLogging.CreateEntry "Waiting Service to start:" & strState, LogTypeInfo
					WScript.Sleep 1000
				Next
				Set colSvcCheck = Nothing
			Loop	
		Next
	
	End Sub

    Sub CopyIniFile(sSrc,sDest,fileformat)

    	Dim line
    	Dim fdst
    	Dim fsrc
        Dim sFileFormat    
        oLogging.CreateEntry "Copy Ini From " & sSrc & " to " & sDest, LogTypeInfo
        If fileformat = -1 Then
            sFileFormat ="Unicode"
        Else
            sFileFormat = "ASCII"
        End If
	
        oLogging.CreateEntry "Using " & sFileFormat & " Mode", LogTypeInfo
        
        Set fsrc = oFSO.OpenTextFile( sSrc, 1, False, fileformat)
        Set fdst = oFSO.OpenTextFile( sDest, 2, True, fileformat)
           
    	Do While fsrc.AtEndOfStream = False
    		line = fsrc.ReadLine
            'oLogging.CreateEntry "Line Read: [" & line & "]", LogTypeVerbose
    		line = Trim(line)
            line = oEnvironment.Substitute(line)
            'oLogging.CreateEntry "Line Modified: [" & line & "]", LogTypeVerbose
            fdst.WriteLine line
    	Loop
    	fsrc.Close
        fdst.Close
    	
    End Sub
    

	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'
	' Class Initialization
	'
	
	Private Sub Class_Initialize
	

	End sub

End Class

