# Project Summary
This project is test querying the Zillow API end points
- Additional information can be found [here] (https://www.zillow.com/howto/api/APIOverview.htm)
# How to get comps:
1. Clone the repository.
2. Obtain a ZWSID from the URL above.
3. Create a file called ```ZWSID.text``` with the key
4. From the project directory, in the command line, write ```python getComps.py "address" "zipcode" "# of comps"```
    - ```address``` is the address of the property you want the comparables for
    - ```zipcode``` is the zip code of the property
    - ```# of comps``` is a number between **1** and **25**