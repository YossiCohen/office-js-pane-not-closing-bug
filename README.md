
Simple addin that demonstrate the bug in Excel add-in (Office 365)
------------------------------------------------------------------

Steps to load the addin:
1. Share the root folder of this project and Add as trusted catalog (or copy the mainfest existing catalog)
2. Start the server: python -m http.server 8080
3. load the add-in in excel

About the bug:
If there is a single interval related pane - closing it will stop the interval.
If there is another pane open when closing the pane with interval - it will continue to run until reopening the interval pane 

Steps to reproduce the bug:
there are 3 different scenarions:
scenario A:
1. Open the "interval update pane"
2. Click the "Random interval!" button (see how values change rapidly in the active worksheet)
3. Close the "interval update pane"
Result: Values stop changing rapidly

Senario B:
1. Open the "Highlight pane"
2. Open the "Interval update pane"
3. Click the "Random interval!" button (see how values change rapidly in the active worksheet)
Mid-step-result: Values keep on changing rapidly
4. Close the "Interval update pane"
Mid-step-result: Values keep on changing rapidly
5. Close the "Highlight pane"
Mid-step-result: Values keep on changing rapidly
6. Open the "Highlight pane"
Mid-step-result: Values keep on changing rapidly
7. Close the "Highlight pane"
Mid-step-result: Values keep on changing rapidly
8. Open the "interval update pane"
Result: Values stop changing rapidly

Scenario C:
1. Open the "Highlight pane"
2. Open the "Interval update pane"
3. Click the "Random interval!" button (see how values change rapidly in the active worksheet)
Mid-step-result: Values keep on changing rapidly
4. Close the "Highlight pane"
Mid-step-result: Values keep on changing rapidly
5. Open the "Interval update pane"
Result: Values stop changing rapidly

To summarize that:
If there is a single interval related pane - closing it will stop the interval.
If there is another pane open when closing the pane with interval - it will continue to run until reopening the interval pane 
