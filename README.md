# Reports Generate Tools - Session Folder and Module Download

Only two requested behavior changes are included:

1. A separate `Download <Module> Excel` button remains above the report rows for every selected module, even when every report count is 0. It downloads a newly generated Excel containing only the selected module rows while preserving the master headings and column order.
2. The connected folder, parsed workbooks, and dynamic module list are kept in application memory while the website tab remains open. Moving back to the dashboard and returning to Reports and Status does not require reconnecting the Excel folder. Closing or refreshing the browser tab requires reconnecting for browser-security reasons.

No dashboard, styling, module-count, report-count, Integrated Tool, navigation, or other behavior was changed.
