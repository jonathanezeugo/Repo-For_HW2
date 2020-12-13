# Repo-For_HW2
This HW loops through all stocks data for 3 subsequent year and outputing results that help determine percentage annual change in price for opening to closing for each year. It also presents the total stock volume for each year and comparatively with the other years.


Personal Note:
I had a problem that I tried various ways and even consulted with others but we could not seem to understand what the problem was. Calculating the Total Volume was giving me Runtime Error Code 6 – Overflow. I tried to increase the type from double to long but it made no difference. I had to reorder the positioning of the total stock volume calculation. It kept giving me values 217 billion. I noticed the picture provided us as guide showed that for 2014, we’re supposed to have about 57978100 as Total Stock Volume but when my code tries to calculate if I place the formular at a different position from the current placement, it gives me the error code 6. I had to turn it in, but I know something is wrong with my buffer capacity despite change type to Long from Double.
Therefore, my Total Stock Volume for each year would be very much less than I was supposed to get, but it’s some fault in the computing aggregation. I will like to investigate this further.


UPDATE:
I am turning in my corrected HW2 just in case it can still be considered. I had already done it at about when I turned it in, but I didn't know how I could turn it in if it wasn't requested. I however am submitting it just in case I could be considered for a review of my grade on this HW. I did not use the solution provided though, and this was already set up before the solution was provided. You can also see that my script is very different from provided solution. Thanks for your kind consideration.