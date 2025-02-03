	<SCRIPT TYPE="text/javascript">
		var numb = '0123456789';
		var lwr = 'abcdefghijklmnopqrstuvwxyz';
		var upr = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ';
		var spcl = " !#$%&'()*+,-./:;<=>?@[\]^_`{|}~"
		 
		function isValid(parm,val) 
		{
			if (parm == "") return false;
			for (i=0; i<parm.length; i++) 
				{
					if (val.indexOf(parm.charAt(i),0) == -1) return false;
				}
			return true;
		}
		 
		function countType(parm,val) 
		{
			var counter;
			counter=0;
			if (parm == "") return false;
			for (i=0; i<parm.length; i++) 
				{
					if (val.indexOf(parm.charAt(i),0) > -1) 
					{
						counter++;
					}
				}
			return counter;
		}
		function isNum(parm) {return isValid(parm,numb);}
		function isLower(parm) {return isValid(parm,lwr);}
		function isUpper(parm) {return isValid(parm,upr);}
		function isAlpha(parm) {return isValid(parm,lwr+upr);}
		function isAlphanum(parm) {return isValid(parm,lwr+upr+numb);}
		function isSpecial(parm) {return isValid(parm,spcl);}

		function countNum(parm) {return countType(parm,numb);}
		function countLower(parm) {return countType(parm,lwr);}
		function countUpper(parm) {return countType(parm,upr);}
		function countAlpha(parm) {return countType(parm,lwr+upr);}
		function countAlphanum(parm) {return countType(parm,lwr+upr+numb);}
		function countSymbols(parm) {return (parm.length - countType(parm,lwr+upr+numb));}
		function countSpecial(parm) { return countType(parm, spcl) };
		function countInvalid(parm) { return (parm.length - countType(parm, lwr + upr + numb + spcl)); }
</SCRIPT>