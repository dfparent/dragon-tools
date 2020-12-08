'#Uses "paths.bas"

Public sub runAutoIt(script as string, optional parameters as variant)
		dim commandLine as string
		commandLine = """" & getPath("auto it") & """ """ & getPath("auto it script") & "\" & script & """"
		if not IsMissing(parameters) then
			if VarType(parameters) = vbString then
				commandLine = commandLine & " """ & parameters & """"
			elseif IsArray(parameters) then
				dim parameter as string
				for each parameter in parameters
					'msgbox parameter
					commandLine = commandLine & " """ & parameter & """"
				next
			end if
		end if
		
		'msgbox commandLine
		'Clipboard commandLine
		Shell commandLine
end sub 

