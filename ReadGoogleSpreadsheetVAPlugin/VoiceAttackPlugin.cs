// -----------------------------------------------------------------------
// <copyright file="VoiceAttackPlugin.cs" company="Insequence Corporation">
//      Copyright (c) Insequence Corporation. All Rights Reserved
// </copyright>
// -----------------------------------------------------------------------
namespace ReadGoogleSpreadsheetVAPlugin
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Text;

    public class VoiceAttackPlugin
    {
        public static string VA_DisplayName()
        {
            return "Read Google Spreadsheet Plugin - v1.0.0";
        }

        public static string VA_DisplayInfo()
        {
            var sb = new StringBuilder();
            sb.AppendLine("Read Google Spreadsheet Plugin");
            sb.AppendLine("Author: Jeffrey Geer");
            sb.Append("Reads data from a Google Doc's Spreadsheet");
            return sb.ToString();
        }

        public static Guid VA_Id()
        {
            // {CD39E757-B3B4-4C00-9AA2-8B861FAC391D}
            return new Guid("{CD39E757-B3B4-4C00-9AA2-8B861FAC391D}");
        }

        static Boolean _stopVariableToMonitor = false;

        //this function is called from VoiceAttack when the 'stop all commands' button is pressed or a, 'stop all commands' action is called.  this will help you know if processing needs to stop if you have long-running code
        public static void VA_StopCommand()
        {
            _stopVariableToMonitor = true;
        }


        //note that in this version of the plugin interface, there is only a single dynamic parameter.  All the functionality of the previous parameters can be found in vaProxy
        public static void VA_Init1(dynamic vaProxy)
        {
            //this is where you can set up whatever session information that you want.  this will only be called once on voiceattack load, and it is called asynchronously.
            //the SessionState property is a local copy of the state held on to by VoiceAttack.  In this case, the state will be a dictionary with zero items.  You can add as many items to hold on to as you want.
            //note that in this version, you can get and set the VoiceAttack variables directly.

            /*
            vaProxy.SessionState.Add("new state value", 369);  //set whatever private state information you want to maintain (note this is a dictionary of (string, object)
            vaProxy.SessionState.Add("second new state value", "hello");

            vaProxy.SetSmallInt("initializedCondition1", 1);  //set some meaningless example short integers (used to be called, 'conditions')
            vaProxy.SetSmallInt("initializedCondition2", 2);

            vaProxy.SetText("initializedText1", "This is 1");  //set some meaningless example text values
            vaProxy.SetText("initializedText2", "This is 2");

            vaProxy.SetInt("initializedInt1", 55);  //set some meaningless example integer values
            vaProxy.SetInt("initializedInt2", 77);

            vaProxy.SetDecimal("initializedDecimal1", 44.2m);  //set some meaningless example decimal values
            vaProxy.SetDecimal("initializedDecimal2", -99.9m);

            vaProxy.SetBoolean("initializedBoolean1", true);  //set some meaningless example boolean values
            vaProxy.SetBoolean("initializedBoolean2", false);
            //*/
        }

        public static void VA_Exit1(dynamic vaProxy)
        {
            //this function gets called when VoiceAttack is closing (normally).  You would put your cleanup code in here, but be aware that your code must be robust enough to not absolutely depend on this function being called
            /*
            if (vaProxy.SessionState.ContainsKey("myStateValue"))  //the sessionstate property is a dictionary of (string, object)
            {
                //do some kind of file cleanup or whatever at this point
            }
            //*/
        }

        public static void VA_Invoke1(dynamic vaProxy)
        {
            /*
            //This function is where you will do all of your work.  When VoiceAttack encounters an, 'Execute External Plugin Function' action, the plugin indicated will be called.
            //in previous versions, you were presented with a long list of parameters you could use.  The parameters have been consolidated in to one dynamic, 'vaProxy' parameter.

            //vaProxy.Context - a string that can be anything you want it to be.  this is passed in from the command action.  this was added to allow you to just pass a value into the plugin in a simple fashion (without having to set conditions/text values beforehand).  Convert the string to whatever type you need to.

            //vaProxy.SessionState - all values from the state maintained by VoiceAttack for this plugin.  the state allows you to maintain kind of a, 'session' within VoiceAttack.  this value is not persisted to disk and will be erased on restart. other plugins do not have access to this state (private to the plugin)

            //the SessionState dictionary is the complete state.  you can manipulate it however you want, the whole thing will be copied back and replace what VoiceAttack is holding on to


            //the following get and set the various types of variables in VoiceAttack.  note that any of these are nullable (can be null and can be set to null).  in previous versions of this interface, these were represented by a series of dictionaries

            //vaProxy.SetSmallInt and vaProxy.GetSmallInt - use to access short integer values (used to be called, 'conditions')
            //vaProxy.SetText and vaProxy.GetText - access text variable values
            //vaProxy.SetInt and vaProxy.GetInt - access integer variable values
            //vaProxy.SetDecimal and vaProxy.GetDecimal - access decimal variable values
            //vaProxy.SetBoolean and vaProxy.GetBoolean - access boolean variable values
            //vaProxy.SetDate and vaProxy.GetDate - access date/time variable values

            //to indicate to VoiceAttack that you would like a variable removed, simply set it to null.  all variables set here can be used withing VoiceAttack.
            //note that the variables are global (for now) and can be accessed by anyone, so be mindful of that while naming


            //if the, 'Execute External Plugin Function' command action has the, 'wait for return' flag set, VoiceAttack will wait until this function completes so that you may check condition values and
            //have VoiceAttack react accordingly.  otherwise, VoiceAttack fires and forgets and doesn't hang out for extra processing.


            //below is just some sample code showing how to work with vaProxy.  There's more in the VoiceAttack help documentation that is installed with VoiceAttack (VoiceAttackHelp.pdf).

            if (vaProxy.GetText("myCSharpTestValue") != null) //was the text value passed set?
            {
                vaProxy.SetText("myCSharpTestValue", "hello " + new Random().Next(1, 6).ToString()); //if the value is not null, this is a subsequent call... just change up the value to nonsense.  this value will go back to VoiceAttack (perhaps to be read with TTS or whatever)
            }
            else //value has not been set.  set it here
            {
                vaProxy.SetText("myCSharpTestValue", "this is new");
            }

            short? testShort = vaProxy.GetSmallInt("someValueIWantToClear");  //note that we are using short? (nullable short) in case the value is null
            if (testShort.HasValue)
            {
                vaProxy.SetSmallInt("someValueIWantToClear", null);  //setting the value to null tells VoiceAttack that you want the variable removed
            }

            //here we check the context to see if we should perform an action (with some additional examples of what can be done with vaProxy
            if (vaProxy.Context == "fire weapons")
            {
                if (vaProxy.CommandExists("secret weapon"))  //check if a command exists
                {
                    if (vaProxy.ParseTokens("{ACTIVEWINDOWTITLE}") == "My Awesome Game")  //check the active window title using the parsetokens function
                    {
                        vaProxy.ExecuteCommand("secret weapon");  //this tells VoiceAttack to execute the, 'secret weapon' command by name (if the command exists
                    }
                    else
                    {
                        vaProxy.WriteToLog("Your game was not active and on top.", "purple");
                    }
                }
                else
                {
                    vaProxy.WriteToLog("the secret weapon command does not exist.  you deleted it, didn't you?", "orange");
                }
            }

            //here we are adding some stuff to state
            object objValue = null;
            String strStateValue = null;
            if (vaProxy.SessionState.TryGetValue("myStateValue", out objValue))  //we check to see if there is a value in state for 'myStateValue'
            {
                strStateValue = (String)objValue; //if we find something, we are going to cast it as a string and use it somewhere in here...
            }
            else
                strStateValue = "initialized";  //nothing was found in state... just set the string to, 'initialized' and keep moving...

            //hope that helps some!
            //*/
        }
    }
}
