#
# Regular cron jobs for the mohan-0.13 package
#
0 4	* * *	root	[ -x /usr/bin/mohan-0.13_maintenance ] && /usr/bin/mohan-0.13_maintenance
