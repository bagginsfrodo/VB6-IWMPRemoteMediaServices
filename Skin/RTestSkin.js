/*
Demo Skin JScript
©2015 Kevin Lincecum AKA FrodoBaggins   email: baggins DOT frodo AT_SYMBOL gmail DOT com
License: Free usage as long as you send me an email and mention me somewhere in your readme, about, etc

Skin Programming Reference
https://msdn.microsoft.com/en-us/library/windows/desktop/dd564359%28v=vs.85%29.aspx
https://msdn.microsoft.com/en-us/library/windows/desktop/dd564952%28v=vs.85%29.aspx
*/




var Testing;

function OnLoad()
{
    MyScriptableObject.InitScript(this);

    //MyScriptableObject.alert('Howdy');

    Testing = MyScriptableObject.GetTestObj();

    MyScriptableObject.Woot();
};




function Event_MyScriptableObject(msg)
{
    MyScriptableObject.alert(msg);
};


function Event_TestObj(msg)
{
    Testing.alert(msg);
};


function Event_DirectScriptAccess(msg)
{
    MyScriptableObject.alert(msg);
};



function YFM()
{
    MyScriptableObject.alert('You Found Me!');
};