The algorithm is divided into two parts to facilitate understanding. The objective is to create an updatable database that is easy to obtain and execute using public data available on the web.

The first part of the algorithm QAP.py
-
- Involves obtaining data from various satellite catalogues that are available from open sources on the internet.

The second part of the algorithm PROCESAMIENTO.PY
-
involves post-processing the data that has already been obtained, intending to verify the number of units for each cubesat and create the final database

First, run QAP.py 
Takes about 8-12 hours
Then run procesamiento.py

Regarding the database buildup care has to be taken with updates in the link structure of the databases
-


General notes
-
-  Since the implementation Celestrack has updated its interface and a simpler, yet to-be-implemented query for the Satellite Catalogue SatCat exists   See https://celestrak.org/satcat/satcat-format.php for more details.
-  The more demanding and time-consuming aspect of the tool was the NSSDC query to update the launch date in decayed spacecraft. No specific measures of the time of each step were taken, yet most of the development-phased crashes were in this phase.

