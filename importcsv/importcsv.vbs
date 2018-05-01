option explicit

class importcsv
	public objFile, delimiter, objHeader, items(), count, max

	private sub class_initialize
		delimiter = ","
		set objFile = nothing
	end sub

	public sub close
		if not objFile is nothing then
			objFile.close
			set objFile = nothing
		end if
	end sub

	private sub class_terminate
		close
	end sub

	private function unquote(quote, part)
		dim chunk, physical, width
		chunk = split(part, """")
		width = ubound(chunk)

		for physical = 1 to width - 1
			if chunk(physical) <> "" then exit for
			chunk(physical) = """"
			physical = physical + 1
		next

		if physical <= width then
			unquote = true
			for physical = physical + 1 to width
				chunk(physical) = """" & chunk(physical)
			next
		else
			unquote = false
		end if

		quote = quote & join(chunk, "")
	end function

	public function open(file, header)
		close

		if vartype(file) = 8 then
			set objFile = createobject("Scripting.FileSystemObject").opentextfile(file, 1)
			if err.number <> 0 then
				open = false
				exit function
			end if
		else
			set objFile = file
		end if

		set objHeader = createobject("Scripting.Dictionary")
		dim logical, chunk, physical, quote, width

		if not isnull(header) then
			width = ubound(header)
			for logical = 0 to width
				if not objHeader.exists(header(logical)) then objHeader(header(logical)) = logical
			next
		elseif objFile.atendofstream <> true then
			logical = 0

			chunk = split(objFile.readline, delimiter)
			width = ubound(chunk)
			physical = 0

			do while physical <= width
				if left(chunk(physical), 1) = """" then
					quote = ""
					if not unquote(quote, mid(chunk(physical), 2)) then
						do while true
							physical = physical + 1
							if physical <= width then
								quote = quote & delimiter
								if unquote(quote, chunk(physical)) then exit do
							else
								quote = quote & vbCrLf
								if objFile.atendofstream = true then exit do
								chunk = split(objFile.readline, delimiter)
								width = ubound(chunk)
								physical = 0

								if 0 <= width then if unquote(quote, chunk(0)) then exit do
							end if
						loop
					end if
					quote = ucase(quote)
				else
					quote = ucase(chunk(physical))
				end if

				if not objHeader.exists(quote) then objHeader(quote) = logical
				logical = logical + 1
				physical = physical + 1
			loop
		end if

		if objHeader.count = 0 then
			close
			open = false
			exit function
		end if

		max = logical - 1
		redim items(max)
		open = true
	end function

	public function readrecord
		if objFile.atendofstream = true then
			readrecord = false
			exit function
		end if

		dim logical, chunk, physical, physical2, quote, width, continue
		logical = 0

		chunk = split(objFile.readline, delimiter)
		width = ubound(chunk)
		physical = 0

		do
			for physical = physical to width
				if left(chunk(physical), 1) = """" then
					quote = ""
					if not unquote(quote, mid(chunk(physical), 2)) then
						continue = true
						do while true
							for physical2 = physical + 1 to width
								quote = quote & delimiter
								if unquote(quote, chunk(physical2)) then exit for
							next

							if physical2 <= width then
								physical = physical2
								exit do
							end if

							quote = quote & vbCrLf

							if objFile.atendofstream = true then
								physical = width
								exit do
							end if

							continue = false

							chunk = split(objFile.readline, delimiter)
							width = ubound(chunk)
							physical = 0

							if 0 <= width then if unquote(quote, chunk(0)) then exit do
						loop
					end if

					if logical <= max then
						items(logical) = quote
						logical = logical + 1
					end if

					if not continue then
						physical = physical + 1
						exit for
					end if
				else
					if logical <= max then
						items(logical) = chunk(physical)
						logical = logical + 1
					end if
				end if
			next
		loop while physical <= width

		count = logical
		for logical = logical to max
			items(logical) = null
		next

		readrecord = true
	end function

	public function item(title)
		title = ucase(title)
		if objHeader.exists(title) then item = items(objHeader(title)) else item = null
	end function
end class
