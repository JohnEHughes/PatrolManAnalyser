# PatrolManAnalyser
<p>A simple program that analyses a patrolman report and displays patrols completed based on certain criteria.

At my day job, my security team have a PatrolMan system that involves a NFC device being swiped on tags at locations around a building to prove they have been to those locations. The software the system produces is very rudimentary and amounts to a csv file showing a long list of points. 
After an intensive audit, this was shown to be inadequate and I decided to put together a program to parse the csv report and immediatley display all the points hit, how many completed patrols and more importantly, any points that had been missed. This was a great test of my Python skills and allowed me to practice methods, GUI using Tkinter, XslxWriter to run an excel report as well as a lot of the programming fundamentals such as loops, conditionals, variables and operators.

The difficult part was filtering the data to calculate points hit on certain patrols. There are three different patrols, some points overlap. The patrols are completed at different times in the day. Important factor is adding only one unique point each time per patrol, otherwise an officer could swipe one point many times, and the total points for th eindividual patrol would show complete, but the patrol was not actually completed by hitting all points.

The program displays each patrol separatley in different lable frames as well as a summary section. The listbox is used to display any points missed on each patrol. The point names are reformatted to display a clean name, appended with which patrol and then ordered sequentially to provide the user a quicker understanding.

I have added a 'Run Report' button which produces an excel report detailing the summary data and all points hit. On the second tab it displays which points have been missed with a section for the user to comment a reason and then another box for them to sign and date. This is then archived on the site central storage to provide a snapshot of the patrols completed every day.

To go along with the main program file are three csv files that house the master point sheets for each patrol that the python program compares against.

The program has the 
</p>

<b>Concepts Covered</b>
 <ol>
  <li>Operators</li>
  <li>Scope</li>
  <li>String Formatting/Manipulation</li>
  <li>Variables</li>
  <li>Loops</li>
  <li>Conditionals</li>
  <li>Methods/Arguments</li>
  <li>External Libraries:
    <ul>
     <li>Date Time</li>
     <li>XlsxWriter</li>
     <li>csv</li>
     <li>Tkinter GUI</li>
   </ul>
  </li>
</ol>
