on executeOllama(shellCommand)
	try
		return do shell script shellCommand
	on error errMsg
		return "Error: " & errMsg
	end try
end executeOllama
