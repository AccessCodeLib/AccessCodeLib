BEGIN{
	insideAttributes=0;
	insideComment=0;
	insideModule=0;
	insideClass=0;
	insideType=0;
	insideMethod=0;
	wasClassComment=1; #geht noch nicht.. deswegen so
	insideType=0;
	insideEnum=0;
	firstComment=1;
	fullLine=1;
	lastLine="";
	implements="";
	lastValue="";
	name="";
	classPrinted=0;
	insideFile=1;
	insideModule=0;
	fileComment=0;
}

# Problemumgehung mit "'" umgehen: (nach /#commentare entfernen/ wieder auf "'" stellen)
/.*"'"/ {
	gsub("\"'\"","\"@@@---DASISTEINHOCHKOMMA---@@@\"");
}

#Zeilenumbrüche entfernen
fullLine==0{
	fullLine=1;
	$0= lastLine$0;
	lastLine="";
}
/_$/{
	fullLine=0;
 	sub(/_$/,"");
 	lastLine=$0;
 	#lastLine=substr($0,1,index($0,"_")-1);
 	next;
}

#funktionsenden erkennen
/^[[:blank:]]*End[[:blank:]]+Sub/ && insideMethod==1{
	insideMethod=0;
}

/^[[:blank:]]*End[[:blank:]]+Function/ && insideMethod==1{
	insideMethod=0;
}

#code kommentare erkennen

/^[[:blank:]]*''/ && insideComment!=1 && insideMethod!=1{
	sub("[[:blank:]]*''","///");
	print $0;
        next;
}

/^[[:blank:]]*'[/][/]/ && insideComment!=1 && insideMethod!=1{
	sub("[[:blank:]]*'[/][/]","'///");
	insideComment=1;
}

/^[[:blank:]]*'[[:blank:]]*[/][*][*]/ && insideComment!=1 && insideMethod!=1{
	insideComment=1;
}

#!(/[[:blank:]]*'/) && insideComment==1{
#	$0="'**/"$0
#}

/^[[:blank:]]*'/ && insideMethod!=1{
	if(insideComment==1){
		commentString=substr($0,index($0,"'")+1);
		if(insideFile==1 && fileComment==0)
		{
			fileComment=1;
			commentString=commentString"\n@file "FILENAME"\n";
		}
		print commentString;
	}
}

/[[:blank:]]*Implements[[:blank:]]+/ && insideMethod!=1{
	sub("[[:blank:]]*Implements[[:blank:]]+","");
	implements= ": public "$0;
	sub(".*","");
}

/[[:blank:]]*'[^@]*@class/ && insideMethod!=1{
	if(insideComment==1)	wasClassComment = 1;
}

(/[[:blank:]]*'[*]*[*][/]/ || !(/[[:blank:]]*'/)) && insideComment==1 && insideMethod!=1{
	insideComment=0;
	if(firstComment==1)
	{
		firstComment= 0;
		if( (insideClass||insideModule) && classPrinted == 0 )
		{
			$0="";
			classPrinted= 1;
			if( insideClass )
			{
				print "class "name" "implements"{";
			}
			else if( insideModule )
			{
				print "namespace "name"{";
			}
		}
	}
}

#!/[[:blank:]]*'/ && insideComment==1{
#	insideComment=0;
#	if(firstComment==1)
#	{
#		firstComment= 0
#		if( insideClass && classPrinted == 0 )
#		{
#			$0=""
#			classPrinted= 1;
#			print "class "name"{";
#		}
#	}
#}

#commentare entfernen
/.*'/ && insideMethod!=1{
	gsub("'.*$","");
}

# Prolemumgehung wieder zurücksetzen
/.*"@@@---DASISTEINHOCHKOMMA---@@@"/ {
	gsub("\"@@@---DASISTEINHOCHKOMMA---@@@\"","\"'\"");
}


/[[:blank:]]*End[[:blank:]]+Enum/ && insideEnum==1 && insideMethod!=1{
	print "};";
	insideEnum=0;
}

/[[:blank:]]*End[[:blank:]]+Type/ && insideType==1 && insideMethod!=1{
	print "};";
	insideType=0;
}

insideType==1 && insideMethod!=1 && insideComment!=1{
	if( gsub(/\(\)[[:blank:]]+As[[:blank:]]+/," [() As ") > 0 )	$0= $0"]";
	else if( gsub("[[:blank:]]+As[[:blank:]]+"," [As ") > 0 )	$0= $0"]";
	print $0";";
}

insideEnum==1 && insideMethod!=1 && insideComment!=1{
	print $0",";
}

#Enum erkennen
/^Enum[[:blank:]]+/ || /[[:blank:]]+Enum[[:blank:]]+/ && insideMethod!=1{
	structStart=$0;
	#at line-start
  sub("^Private[[:blank:]]+","private: ",structStart);
	sub("^Public[[:blank:]]+","public: ",structStart);
	sub("^Friend[[:blank:]]+","friend ",structStart);
	sub("^Enum[[:blank:]]+","enum ",structStart);
  #or as whole word
  sub("[[:blank:]]+Private[[:blank:]]+","private: ",structStart);
	sub("[[:blank:]]+Public[[:blank:]]+","public: ",structStart);
	sub("[[:blank:]]+Friend[[:blank:]]+","friend ",structStart);
	sub("[[:blank:]]+Enum[[:blank:]]+","enum ",structStart);
	print structStart" {";
	insideEnum=1;
	next;
}

#Type erkennen
/^Type[[:blank:]]+/ || /[[:blank:]]+Type[[:blank:]]+/ && insideMethod!=1{
	structStart=$0;
	#at line-start
  sub("^Private[[:blank:]]+","private: ",structStart);
	sub("^Public[[:blank:]]+","public: ",structStart);
	sub("^Friend[[:blank:]]+","friend ",structStart);
	sub("^Type[[:blank:]]+","struct ",structStart);
  #or as whole word
  sub("[[:blank:]]+Private[[:blank:]]+","private: ",structStart);
	sub("[[:blank:]]+Public[[:blank:]]+","public: ",structStart);
	sub("[[:blank:]]+Friend[[:blank:]]+","friend ",structStart);
	sub("[[:blank:]]+Type[[:blank:]]+","struct ",structStart);
	print structStart" {";
	insideType=1;
}

#Klasse erkennen
/[[:blank:]]*VERSION[^C]+CLASS/ && insideMethod!=1{
	insideClass=1;
	insideModule=0;
}

/[[:blank:]]*Begin[^V]+VB\.UserControl[[:blank:]]+/ && insideMethod!=1{
	insideClass=1;
	insideModule=0;
}

/[[:blank:]]*Begin[^V]+VB\.Form[[:blank:]]+/ && insideMethod!=1{
	insideClass=1;
	insideModule=0;
}

/[[:blank:]]*Attribute[[:blank:]]+VB_Name[[:blank:]]+/ && insideMethod!=1{
	match( $0, /"[^"]*"/ );
	name= substr( $0, RSTART+1, RLENGTH-2 );
	insideAttributes=1;
	insideFile=0;
	insideModule=1;
}

insideAttributes == 1 && !(/[[:blank:]]*Attribute[[:blank:]]+/ || /^[[:blank:]]*$/) && insideMethod!=1{
	insideAttributes=0;
	if( match( $0, /[[:blank:]]*'[[:blank:]]*[/][*][*]/ ) == 0 && classPrinted == 0 )
	{
		if( insideClass )
		{
			print "class "name" "implements"{";
		}
		else if( insideModule )
		{
			print "namespace "name"{";
		}

		classPrinted=1;
	}
}


!(/[[:blank:]]*Function[[:blank:]]/ || /[[:blank:]]*Property[[:blank:]]/ || /[[:blank:]]*Sub[[:blank:]]/ || /[[:blank:]]*Event[[:blank:]]/) && insideMethod!=1{
	if( gsub(/\(\)[[:blank:]]+As[[:blank:]]+/," [() As ") > 0 )	$0= $0"]";
	else if( gsub("[[:blank:]]+As[[:blank:]]+"," [As ") > 0 )	$0= $0"]";
}

fullLine == 1 && insideComment == 0 && insideType==0 && !(/End[[:blank:]]/) && !(/^Type[[:blank:]]+/ || /[[:blank:]]+Type[[:blank:]]+/ ) && (/^[[:blank:]]*Private[[:blank:]]/ || /^[[:blank:]]*Public[[:blank:]]/ || /^[[:blank:]]*Friend[[:blank:]]/ || /^[[:blank:]]*Const[[:blank:]]/ || /^[[:blank:]]*Function[[:blank:]]/ || /^[[:blank:]]*Declare[[:blank:]]/ || /^[[:blank:]]*Property[[:blank:]]/ || /^[[:blank:]]*Sub[[:blank:]]/ || /^[[:blank:]]*Event[[:blank:]]/) && insideMethod!=1{
	# either at the beginning of a line
  sub("^Private[[:blank:]]+","private: ");
	sub("^Public[[:blank:]]+","public: ");
	sub("^Friend[[:blank:]]+","friend ");
  # or between spaces
  sub("[[:blank:]]+Private[[:blank:]]+","private: ");
	sub("[[:blank:]]+Public[[:blank:]]+","public: ");
	sub("[[:blank:]]+Friend[[:blank:]]+","friend ");
	gsub("[[:blank:]]+Lib[[:blank:]]+[^[:blank:]]+","");
	gsub("[[:blank:]]+Alias[[:blank:]]+[^[:blank:]]+","");
	$0=gensub("[[:blank:]]*([a-zA-Z_]+[[:blank:]]*)\\(([0-9]+)[[:blank:]]*\\)","\\1[\\2]","g");
	gsub("[[:blank:]]+"," ");
	if( $0 != "" ){
		print $0";";
	}
}

#funktionen erkennen
/^Sub[[:blank:]]+/ || /[[:blank:]]+Sub[[:blank:]]+/ && insideMethod!=1{
	insideMethod=1;
}

/^Function[[:blank:]]+/ || /[[:blank:]]+Function[[:blank:]]+/ && insideMethod!=1{
	insideMethod=1;
}

/^Declare[[:blank:]]+/ || /[[:blank:]]+Declare[[:blank:]]+/ && insideMethod==1{
	insideMethod=0;
}

# geht nicht deswegen raus
#fullLine == 1 && insideComment == 0 && /End[[:blank:]]Type/{
#	if( lastValue!= "" ) print lastValue;
#	print "};";
#	insideType=0;
#}
#
#insideType == 1{
#	if( lastValue!= "" ) print lastValue",";
#	gsub("[[:blank:]]+As[[:blank:]]+",":");
#	lastValue= $0;
#}
#
#fullLine == 1 && insideComment == 0 && /[[:blank:]]Type[[:blank:]]/{
#	gsub("Type","struct");
#	if( $0 != "" ) print $0"{";
#	insideType=1;
#}

END{
	if(insideClass == 1 || insideModule == 1)
	{
		print "};";
	}
}
