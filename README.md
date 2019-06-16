# localconf
This is an SQLite-based key-value storage mechanism for storing local configuration settings.\
\
**The "key" is actually not a single field but a combination of the fields:**
* application (if you are storing settings not for a signle app but a suite of different apps)
* user (if you are have a central configuration file for all users and not different config files at each user's home folder)
* section (a grouping of settings)
* key (a single configuration option)

**"value" is also not just one field:**
* value (the value for this configuration option)
* comment (some explanatory info for the particular setting, ie accepted/boundary values)

Note that you can store unique application/user/section/key values or you can create an array of such values (all having the same application/user/section/key combination)

There is no extensive documentation, please study the demo application: it should give you a clear idea on how to use localconf.
