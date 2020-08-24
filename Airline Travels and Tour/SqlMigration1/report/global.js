var cur = null; 
var curClass=null;

var selected=new Array();
var selectedClasses=new Array();

// return short file name
function getShortFileName()
{
    // get last position of slash
    var positionOfLastSlash = location.href.lastIndexOf('/') + 1;
    // get last position of dot
    var positionOfLastDot = location.href.lastIndexOf('.');
    // return substring
    return location.href.substring(positionOfLastSlash, positionOfLastDot);
}

// return true is source
function getIsSource()
{
    if (getShortFileName()=="src")
        return true;
    else
        return false;
}

// return true is target
function getIsTarget()
{
    if (getShortFileName()=="trg")
        return true;
    else
        return false;
}


// get current Button frame
function getCurrentButton()
{

    var currentButtonFrameName;
    
    if (getIsSource())
    {
        currentButtonFrameName = "srcButt";
    }
    else if (getIsTarget())
    {
        currentButtonFrameName = "trgButt";
    }
    
    if (currentButtonFrameName != undefined)
    {
        var fr = top.frames['node'];

        var srcButton = null;

        if(fr == undefined)
        {
            frameButton=top.frames[currentButtonFrameName];
        }
        else
        {
    	    frameButton=fr.frames[currentButtonFrameName];
        }

        return frameButton;
    }
}

// create instance buttons manager for bookmarks
function BookmarkButtonsManager()
{
   // return element previous_bookmark
   this.getPreviousButton = function()
   {
        if (getCurrentButton() != undefined)
           return getCurrentButton().document.getElementById("previous_bookmark");
        else
           return undefined;
   }
   
   // return element next_bookmark
   this.getNextButton = function()
   {
        if (getCurrentButton() != undefined)
           return getCurrentButton().document.getElementById("next_bookmark");
        else
           return undefined;
   }
   
   // return element clear_bookmarks
   this.getClearButton = function()
   {
        if (getCurrentButton() != undefined)
            return getCurrentButton().document.getElementById("clear_bookmarks");
        else
            return undefined;
        
   }

   this.disabledButton = function(button)
   {
        if (button != undefined)
        {
            button.onmouseout();
            button.disabled = true;
            button.style.filter = "progid:DXImageTransform.Microsoft.BasicImage(grayscale=1)";
        }
   }
   
   this.enabledButton =  function(button)
   {
        if (button != undefined)
        {
            button.disabled = false;
            button.style.filter = "";
        }
   }

   
}

// constructor of Bookmarks class
function CreateBookmarksObject(Bookmarks)
{
   
    // array of Bookmarks
    var bookmarks = Bookmarks;

    // current position
    var curBookmark = 0;
   
    // get Bookmarks
    this.getBookmarks = function()
    {
        return bookmarks;
    }

    // get current position
    this.getCurrentPosition = function()
    {
       return curBookmark;
    }
    
    // return length of bookmarks
    this.getLength = function()
    {
        return bookmarks.length;
    }
    
    // set current position
    this.setCurrentPosition = function(Position)
    {
        curBookmark = Position;
	    this.setButtonStyle();
    }
    
    // increment position
    this.incPosition = function()
    {
        curBookmark++;
	    this.setButtonStyle();
    }
    
    // decrement position
    this.decPosition = function()
    {
        curBookmark--;
	    this.setButtonStyle();
    }

    // return Bookmark of current position
    this.getCurrentBookmark = function()
    {
	    return bookmarks[curBookmark];
    }
    
    // return Bookmark of position
    this.getBookmarkInPosition = function(Position)
    {
        return bookmarks[Position];
    }

    // disabled/enable buttons
    this.setButtonStyle = function()
    {
        var manager = new BookmarkButtonsManager();
        
        switch (this.getLength())
        {
            case 0: 
            {
		        // if length of bookmarks equal zero
		        // disable all buttons
		        manager.disabledButton(manager.getPreviousButton());
                manager.disabledButton(manager.getNextButton());
                manager.disabledButton(manager.getClearButton());
                break;
            }
            case 1:
            {
		        // if length of bookmarks equal one
		        // disable previous button
		        manager.disabledButton(manager.getPreviousButton());
                manager.enabledButton(manager.getNextButton());
                manager.enabledButton(manager.getClearButton());
                break;
            }
            default :
            {
		        if (this.getCurrentPosition() == 0)
		        {
		            // if current bookmark first in array
		            // disable previous button
		            manager.disabledButton(manager.getPreviousButton());
          	        manager.enabledButton(manager.getNextButton());
          	        manager.enabledButton(manager.getClearButton());
		        }
		        else if (this.getCurrentPosition() == this.getLength()-1)
		        {
			        // if current bookmark last in array
			        // disable next button
			        manager.enabledButton(manager.getPreviousButton());
           	        manager.disabledButton(manager.getNextButton());
                    manager.enabledButton(manager.getClearButton());
		        }
		        else
		        {
			        // if current bookmark not first and not last in array
			        // enabled all button
			        manager.enabledButton(manager.getPreviousButton());
        	        manager.enabledButton(manager.getNextButton());
                    manager.enabledButton(manager.getClearButton());              
		        }
                break;
            }
        }

        
    }   

    this.setButtonStyle(); 
}

// create instance
var Bookmark = new CreateBookmarksObject(new Array());

// create instance buttons manager for problems
function ProblemButtonsManager()
{

   // return element previous_proublem
   this.getPreviousButton = function()
   {
        if (getCurrentButton() != undefined)
            return getCurrentButton().document.getElementById("previous_problem");
        else
            return undefined;
   }
   
   // return element next_proublem
   this.getNextButton = function()
   {
        if (getCurrentButton() != undefined)
            return getCurrentButton().document.getElementById("next_problem");
        else
            return undefined;
   }
   
   // diasabled button
   this.disabledButton = function(button)
   {
        if (button != undefined)
        {
            button.onmouseout();
            button.disabled = true;
            button.style.filter = "progid:DXImageTransform.Microsoft.BasicImage(grayscale=1)";
        }
   }
   
   // enabled button
   this.enabledButton =  function(button)
   {
        if (button != undefined)
        {
            button.disabled = false;
            button.style.filter = "";
        }
   }
   
}


function CreateProblemsObject(Problems)
{
    var problems = Problems
    
    var curProblem = -1;
    
    // return length of problem
    this.getLength = function()
    {
        return problems.length;
    }
    
    // added problem in problems array
    this.addProblem = function(Problem)
    {
        problems[this.getLength()] = Problem;
	    this.setButtonStyle();
    }

    // get current position
    this.getPosition = function()
    {
        return curProblem;
    }

    // set current position
    this.setCurrentPosition = function(Position)
    {
        curProblem = Position;
	    this.setButtonStyle();
    }
    
    // increment position
    this.incPosition = function()
    {
        curProblem++;
	    this.setButtonStyle();
    }
    
    // decrement position
    this.decPosition = function()
    {
        curProblem--;
	    this.setButtonStyle();
    }

    // return Problems of current position
    this.getCurrentProblem = function()
    {
	    return problems[curProblem];
    }
    
    // return Bookmark of position
    this.getProblemInPosition = function(Position)
    {
         return problems[Position];
    }

    // disabled/enable buttons
    this.setButtonStyle = function()
    {
	    var manager = new ProblemButtonsManager();
        
        switch (this.getLength())
        {
            case 0: 
            {
		        // if length of problems equal zero
		        // disable all buttons
		        manager.disabledButton(manager.getPreviousButton());
                manager.disabledButton(manager.getNextButton());
                break;
            }
            case 1:
            {
		        // if length of problems equal one
		        // disable previous button
		        manager.disabledButton(manager.getPreviousButton());
                manager.enabledButton(manager.getNextButton());
                break;
            }
            default :
            {
		        if (this.getPosition() == 0 || this.getPosition() == -1)
		        {
			        // if current problems first in array
			        // disable previous button
			        manager.disabledButton(manager.getPreviousButton());
                    manager.enabledButton(manager.getNextButton());
		        }
		        else if (this.getPosition() == this.getLength()-1)
		        {
			        // if current problems last in array
			        // disable next button
			        manager.enabledButton(manager.getPreviousButton());
                    manager.disabledButton(manager.getNextButton());
		        }
		        else
		        {
			        // if current problems not first and not last in array
			        // enabled all button
			        manager.enabledButton(manager.getPreviousButton());
                    manager.enabledButton(manager.getNextButton());              
		        }

                break;
            }
        }
    }

    this.setButtonStyle();

}

// create instance of Problems
var Problem = new CreateProblemsObject(new Array());

var selectedLevel;

var insideRefl = false;

function oc( s) 
{  

    if (s.substring(0, 1) == "m")  
    {
	var i = document.getElementById(s);
      if(i != null){
	var par = i.parentElement;
        oc(par.id);
        par.scrollIntoView(true);
	return;
      }
    }
	
	
      


	highlightNode(s);
	if(window.event != null)
        window.event.cancelBubble = true;
}

function checkInputParams()
{
  INPUT_PARAMS = new Object();

  separator = ',';

  query = '' + this.location;
  qu = query
  query = query.substring((query.indexOf('?')) + 1);
  if (query.length < 1) { return false; }  
  keypairs = new Object();
  numKP = 1;
  while (query.indexOf('&') > -1) {
    keypairs[numKP] = query.substring(0,query.indexOf('&'));
    query = query.substring((query.indexOf('&')) + 1);
    numKP++;
  }
  keypairs[numKP] = query;

  for (i in keypairs) {
    keyName = keypairs[i].substring(0,keypairs[i].indexOf('='));

    keyValue = keypairs[i].substring((keypairs[i].indexOf('=')) + 1);

    while (keyValue.indexOf('+') > -1) {
      keyValue = keyValue.substring(0,keyValue.indexOf('+')) + ' ' + keyValue.substring(keyValue.indexOf('+') + 1);

    }
    keyValue = unescape(keyValue);

    if (INPUT_PARAMS[keyName]) {
      INPUT_PARAMS[keyName] = INPUT_PARAMS[keyName] + separator + keyValue;

    } else {
      INPUT_PARAMS[keyName] = keyValue;
        // Normal case: name gets value.
    }
  }
  return FORM_DATA;
}

function refl( s) 
{
//alert(s);

var fr= top.frames['node'];
if(fr==undefined)
	fr=top.document;

  
	if(self==top)
		return;
	var obj = null;
	if(self.name=='src')
		obj = fr.frames['trg'];
	else
		obj = fr.frames['src'];
	insideRefl = true;

	obj.highlightAndScrollToVisibleNode( s);	
	//is needed to prevent some IE bugs
	obj.highlightAndScrollToVisibleNode( s);
	
        insideRefl = false;
}



function myscrollIntoView( obj) 
{  



//	if(isinView(obj)>0)
	obj.scrollIntoView(true);

}



function ltrim(argvalue) {

  while (1) {
    if (argvalue.substring(0, 1) != " ")
      break;
    argvalue = argvalue.substring(1, argvalue.length);
  }

  return argvalue;
}

function highlightNode(s)
{
	if(s.length == 0)
		return;

//var spl = s.split(',');
var spl = s.split('||');

cur = document.getElementById(spl[0]);

var i =0;
        for(i = 0; i < selected.length; i++) {
	        var ncur = selected[i];
		ncur.className=selectedClasses[i] ;
		if(ncur.className.indexOf('Div')>0)
			inverseLines(ncur );
	}

        selected=new Array();
	selectedClasses=new Array();


	for(i = 0; i < spl.length; i++) {
		var q = ltrim(spl[i]);

		var temp = q.indexOf("a");
		if((temp == -1 || temp >= 0) && (q.length == 36))
		{
			q = 'a'+q;
		}
	        var ncur1 = document.getElementById(q);
		if(ncur1 != null)
		{
			if(cur == null)
				cur = ncur1;
			selected.push(ncur1);
			selectedClasses.push(ncur1.className);
			if(ncur1.className.indexOf('Div')>0 ||ncur1.className.indexOf('Err')>0 ){
				inverseLines(ncur1 );
				ncur1.className='selectedDiv' ;
			}
			else 
				ncur1.className='selectedSpan' ;

		}



	}
	

}


function inverseLines(s)
{
/*
var oObject = s.all;
if (oObject != null){
   if (oObject.length != null){
      for (i = 0; i < oObject.length; i++){
         var ob = oObject(i);
	 if(ob.className=='line'){
//		 alert(ob.className);
		ob.className ='uls';
	 }
	 else
	 if(ob.className=='uls')
		ob.className ='line';
      }
   }
} 
*/

}


function highlightAndScrollToVisibleNode(s)
{
	highlightNode(s);
        if(cur!= null)
	cur.scrollIntoView(true);
}


function goRef( s) 
{ 
	oc(s); 
	if(cur != null) 
		cur.scrollIntoView(true);
}

function goHref( s) 
{ 
	window.event.cancelBubble = true;
	top.navigate(s); 
}

function toggleBookmark(s) 
{
	if(s.id != "bookmark")
	{
//		s.className = "bookmark";
		s.childNodes[0].src="../bookmark.gif";

		s.id = "bookmark";
	}
	else
	{
//		s.className = "line";
		s.id = "line";
		s.childNodes[0].src="../img/empty.gif";
	}		
	
	Bookmark = new CreateBookmarksObject(document.getElementsByName("bookmark"));
	
	if(Bookmark.getCurrentPosition() == Bookmark.getLength() && Bookmark.getCurrentPosition() > 0)
	{
		Bookmark.decPosition();
	}        
	window.event.cancelBubble = true;
	SetClearBookMarks1();	

		
}

function nextBookmark() 
{
    


	if(Bookmark.getLength() == 0)
		return;
	if(Bookmark.getCurrentPosition() < Bookmark.getLength()-1)
	{
		Bookmark.incPosition();
	}
		goBookmark();
}
function prevBookmark() 
{
    try
    {	
	if(Bookmark.getLength() == 0)
		return;
	if(Bookmark.getCurrentPosition() >0)
	{
		Bookmark.decPosition();
	}
		goBookmark();
    }
    catch(e)
    {

    }

}


function goBookmark() 
{
	
    myscrollIntoView(Bookmark.getCurrentBookmark());
    Bookmark.getCurrentBookmark().scrollIntoView(true);
}



function prevSibling() 
{
	if(cur == null)
		return;
	var par = cur.parentNode;
	if(par!= null)
	{
		var k = 0;
		var siblings = par.children.tags("div");
		for (i = 0; i < siblings.length; i++){
         		if(siblings[i] == cur) {
				k = i;
				break;
				
			}
      		}
		if(k == 0)
			return;

//		highlightAndScrollToVisibleNode(siblings[k-1].id);
siblings[k-1].click();

	}

}
function nextSibling() 
{
	if(cur == null)
		return;
	var par = cur.parentNode;
	if(par!= null)
	{
		var k = 0;
		var siblings = par.children.tags("div");
		for (i = 0; i < siblings.length; i++){
         		if(siblings[i] == cur) {
				k = i;
				break;
				
			}
      		}
		if(k == 0)
			return;

		highlightAndScrollToVisibleNode(siblings[k-1].id);

	}

}
function clearBookmarks() 
{
		var len = Bookmark.getLength();

		for (p = len-1; p >=0 ; p--)
		{
			var k = p;
			
//         		bookmarks[k ].className = "line";
//alert(bookmarks[k ]);
			Bookmark.getBookmarkInPosition(k).childNodes[0].src="../img/empty.gif";
			Bookmark.getBookmarkInPosition(k).id = "line";
			

      		}
		Bookmark.setCurrentPosition(0);
		Bookmark = new CreateBookmarksObject(new Array);
		SetClearBookMarks1();


}

function SetClearBookMarks1() 
{
	try
	{

		if(window.external == null) 
			return;
	
		if(Bookmark.getLength() == 0){
			window.external.DisableClearBookMarks();}
		else
			{window.external.EnableClearBookMarks();}
	}
	catch(e)
	{
	}


}


function nextProblem() 
{

    
    var t = top.frames['messages'];
    if(t!=undefined)
		selectedLevel = t.selectedLev;
	else
		selectedLevel = 0;

	if (Problem.getLength() == 0)
		return;

	var curr = Problem.getPosition();
	
	while(Problem.getPosition() < Problem.getLength() - 1)
	{
	    Problem.incPosition();
	    if(parseInt(Problem.getCurrentProblem().attributes['category'].value, 10) > selectedLevel)
	        continue;
	    goProblem();
	    return;
	}
	
	Problem.setCurrentPosition(curr);
	
}
function prevProblem() 
{
    var t = top.frames['messages'];
    if(t!=undefined)
		selectedLevel = t.selectedLev;
	else
		selectedLevel = 0;
		
	if (Problem.getLength() == 0)
		return;
		
	var curr = Problem.getPosition();
	
	while(Problem.getPosition() > 0)
	{
	    Problem.decPosition();
	    if(parseInt(Problem.getCurrentProblem().attributes['category'].value, 10) > selectedLevel)
	        continue;
	    goProblem();
	    return;
	}
	
	Problem.setCurrentPosition(curr);
	
}

function goProblem() 
{
      
    if(Problem.getCurrentProblem().tagName == "ERR")
        Problem.getCurrentProblem().parentElement.scrollIntoView(true);
    else
        Problem.getCurrentProblem().scrollIntoView(true);
    
	Problem.getCurrentProblem().click();
/*
	var i = problems[curProblem].parentElement;

	if (i.id != null)
		highlightNode(i.id);
*/

}

function addProblem(prb) 
{
    Problem.addProblem(prb);
}


function collectProblems() 
{
		var cands= document.getElementsByTagName("ERR");
		for (i = 0; i < cands.length; i++)
		    {
		    
				    if(cands[i].attributes["category"] != null)
					    addProblem(cands[i]);
    
          
      		}	


}


function isinView(oObject)
{
//    var p = oObject.offsetParent; 
    var p = self.document.body; 
    var t = oObject.offsetTop;
    var h = p.clientHeight;
    var b = oObject.offsetHeight;

	var st = p.scrollTop;

//alert(p);

//alert("(t > h+st || t< st || t+b>h+st )\nt = "+t+"  h = "+h+"  st = "+st+"  b = "+b);
    if (t > h+st || t< st || t+b>h+st ) {
        return 1;
    }
	return 0;
}













