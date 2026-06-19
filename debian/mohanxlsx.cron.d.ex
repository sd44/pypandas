#
# Regular cron jobs for the mohanxlsx package.
#
0 4	* * *	root	[ -x /usr/bin/mohanxlsx_maintenance ] && /usr/bin/mohanxlsx_maintenance
