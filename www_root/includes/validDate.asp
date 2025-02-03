	<SCRIPT type="text/javascript">
	function validDate(fld) 
	{
		var testMo, testDay, testYr, inpMo, inpDay, inpYr, msg;
		var inp = fld.value;
		if (inp=="") return true;
		var testDate = new Date(inp);
		// extract pieces from date object
		testMo = testDate.getMonth() + 1;
		testDay = testDate.getDate();
		testYr = testDate.getFullYear();
		// extract components of input data
		inpMo = parseInt(inp.substring(0, inp.indexOf("/")), 10);
		inpDay = parseInt(inp.substring((inp.indexOf("/") + 1), inp.lastIndexOf("/")), 10);
		inpYr = parseInt(inp.substring((inp.lastIndexOf("/") + 1), inp.length), 10);
		// make sure parseInt() succeeded on input components
		if (isNaN(inpMo) || isNaN(inpDay) || isNaN(inpYr)) 
		{
			msg = "Invalid Date. Enter using form mm/dd/yyyy";
		}
		// make sure that year is reasonable, not just valid!
		if (parseInt(inpYr) < 1980 || parseInt(inpYr) > 2050)
		{
			msg = "Check the year. The value does not seem reasonable.";
		}
		// make sure conversion to date object succeeded
		if (isNaN(testMo) || isNaN(testDay) || isNaN(testYr)) 
		{
			msg = "Couldn't convert your entry to a valid date. Enter using form mm/dd/yyyy. Try again.";
		}
		// make sure values match
		if (testMo != inpMo || testDay != inpDay || testYr != inpYr) 
		{
			msg = "Check the range of your date value. Enter using form mm/dd/yyyy";
		}
		if (msg) 
		{
			// there's a message, so something failed
			alert(msg);
			fld.focus();
			return false;
		} 
		else 
		{
			fld.value = "" + inpMo + "/" + inpDay + "/" + inpYr
			return true;
		}
	}
	</SCRIPT>
