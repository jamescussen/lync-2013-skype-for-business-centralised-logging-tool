Lync 2013 / Skype for Business Centralised Logging Tool
=======================================================

            

The Lync 2013 / Skype for Business Centralised Logging Tool is a GUI that allows you easily control the centralised logging service within Skype for Business On-Prem pools and open log files easily within the Snooper tool. In Skype for Business 2015/2019
 there is now a tool that Microsoft ships with Skype for Business called ClsLogger.exe that can also do this job, however, you may want an alternative, so here it is.

Centralised Logging Tool features:



  *  The tool has the ability to manage logging on multiple Pools at once. When the tool first loads it will learn the current state of all the Pools. The UI can show the state of each pool as they are clicked on in the Pools listbox and only allow the user
 to execute functions that are available in that state (eg. Start and Stop buttons will be disabled if the current state of the pool doesn’t allow you to use them).

  *  When the tool loads it will query the Lync/SfB servers for the status of running logging scenarios, and the tool will display the current state of each pool as it is selected in the pools list. This means that you can use the tool to learn what logging
 is currently running on each pool. 
  *  The scenarios list is dynamic, so Custom Scenarios will also be included in the scenarios list along with the default scenarios.

  *  All pool types, including Persistent Chat pools are listed. If Mediation server is co-located with the front end server it will not be added to the list, but if it’s on a separate physical server it will be listed.

  *  The tool has the ability to have AlwaysOn and Scenario logging enabled at the same time. AlwaysOn logging has its own button and status messages.

  *  The components filter list can either show all the components available for logging on the system, or list only the components for the specific scenario that is currently being logged (or just logged) on a pool.

  *  The Logging duration time can be specified when you Start logging, and running scenarios can have their Duration updated with the new Update button. By default, logging will run for 4 hours and then automatically stop. So if you want to do a long running
 log, or a much shorter log you can specify this when you start logging, or you have the option to Change the duration of a log whilst it’s running. The format of Logging duration is as follows: <days.hours:minutes> (eg. “0.04:00”) .

  *  You can now select either MatchAll, or MatchAny settings. These settings affect the way multiple filters are applied when Searching/Exporting logs. If MatchAll is used, then the filters execute in a logical ‘AND’ fashion, where all the filters
 must match to return results. If MatchAny is selected, then the filters apply in a logical ‘OR’, which means if any of the filters match the result will be returned.

  *  The Call-Id, IP Address, Phone and URI filters have separate text boxes, so these filters can be used in conjunction with each other (and the MatchAll/MatchAny setting) for more precise filtering.

  *  The Components listbox is Multi-Selectable, offering export filtering by one or many component types at once.

  *  Start and End time filter is now included as a filter option. By default the Search command will filter through only the last 30 minutes of logs to return a result. If you need to capture more than this then you can used the Start and End time settings.
 I have made the defaults in these boxes display the range of 1 hour as a starting point. It’s best that you keep the date/time format in the format shown.

  *  The logging folder can be selected using a Folder Browse dialog to having to type the location.

  *  Analyse Log button allows you to choose a log file to open directly in Snooper, to save you having to manually open Snooper and import the logs.


 


For more details please refer to the blog post here:


**[https://www.myskypelab.com/2013/04/lync-2013-centralised-logging-tool.html](https://www.myskypelab.com/2013/04/lync-2013-centralised-logging-tool.html)**


 





        
    
