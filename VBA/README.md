## Overview

This macro stops you accidentally emailing two different clients sensitive information- better safe than sorry!


## Installation

The contents of ThisOutlookSession should be copied in to: 
Outlook -> Developer Tab -> Macros -> Microsoft Outlook Objects -> ThisOutlookSession

Populate the ExampleFile with names, companies or addresses that should never mutually receive an email.
Each entry (either recipient name, or email address) should be on its own line.
 
The contents of CC_Filter can be copied into a new macro module. Please ensure the Filepath variable in the TimerCallback sub points to the location of ExampleFile somewhere on your PC.


## Use Instructions

If installed correctly, the filter should automatically begin when you next start Outlook.
Ensure macros are enabled, and delay your outgoing mail by 2-3 minutes for optimal effectiveness.
