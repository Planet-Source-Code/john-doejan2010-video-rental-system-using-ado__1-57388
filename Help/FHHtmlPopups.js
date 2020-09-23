// Entire contents copyright (C) 1999-2002, Work Write, Inc.
// and KeyWorks Software. All rights reserved.
// Contact: Cheryl Lockett Zubak at cheri@workwrite.com

var KeyPopup;   // KeyHelp.KeyPopup  ActiveX object

function DisplayHtmlPopup(URL,left,top,width)
{
  if (!KeyPopup)
    KeyPopup = new ActiveXObject("KeyHelp.KeyPopup");
  KeyPopup.Width = width;
  KeyPopup.DisplayURL(URL,left,top);
}
