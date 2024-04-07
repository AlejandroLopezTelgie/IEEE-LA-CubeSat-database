CubeSat database algorithm
-
- The algorithm is divided into two parts for easier understanding, the goal being to obtain an updatable database, easy to obtain and execute from public data on the web.

The first part of the algorithm QAP.py
-
- Involves obtaining data from various satellite catalogues that are available from open sources on the internet.
- The first algorithm consists of obtaining data from different satellite catalogues available in open sources on the Internet, such as Celestrack, NASA National Space Science Data Center (NSSDC) and Union of Concerned Scientists, Cospar ID is used to match the information between the different databases and keywords are used to speed up the execution.
- databases and keywords are used to speed up the execution of the software.

The second part of the algorithm PROCESAMIENTO.PY
-
-Involves post-processing the data that has already been obtained, intending to verify the number of units for each cubesat and create the final database

How to use it?
- 
- First, run QAP.py
- Takes about 8-12 hours
- Then run procesamiento.py
- Make your plots in the tool of your choosing (we used Microsoft Excel)

General notes
-
-  Since the implementation Celestrack has updated its interface and a simpler, yet to-be-implemented query for the Satellite Catalogue SatCat exists   See https://celestrak.org/satcat/satcat-format.php for more details.
-  The more demanding and time-consuming aspect of the tool was the NSSDC query to update the launch date in decayed spacecraft. No specific measures of the time of each step were taken, yet most of the development-phased crashes were in this phase.
-  Regarding the database buildup care has to be taken with updates in the link structure of the databases

