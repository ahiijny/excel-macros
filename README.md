// I know this code is unreadable and bad but I wrote these in like 2015 and they still work so whatevs

- yomu.bas: reading tracker VBA script
	- bind PlusPlus() to Ctrl + M
	- for sheet "Finput", assumed existing columns: "Title", "Fandom", "Ch", "Date", "Author", "Last Entry:" (with corresponding number in cell to right)
	- for active sheet, assumed existing columns: "Title", "Fandom", "Ch", "Author", "Link"
- miru.bas: watching tracker VBA script
	- bind PlusPlus() to Ctrl + P
	- for sheet "Episodes", assumed existing columns: "Date", "Title", "S", "Ep", "Subtitle", "Last Entry:" (with corresponding number in cell to right)
	- for active sheet, assumed existing columns: "Studio", "Translation", "Title", "S", "Ep", "Subtitle"
- budget.bas: spending tracker VBA script
	- bind buildsome() to Ctrl + Q
	- bind buildall() to Ctrl + Shift + Q
	- for sheet "STREAM", assumed existing columns: "Date", "Description", "From", "To", "Amount", and countref in H1, statusref in K1
	- for sheet "PAGE", assumed existing columns: "Alias", "Name", "Stream", "Archive"
    