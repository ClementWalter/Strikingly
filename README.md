# Strikingly
Simple CMS for Strikingly simple store in Google app scripts

Theses scripts can be added to a Google Spreadsheet in a bounded project to
have a simple CMS for Strikingly store. It used the `orders.csv` export file
to expand and generate client orders.

Once the orders are well formatted in a regular spreadsheet, the user can then modify
them (usual client requests are shipping address or product) and update their status.
The script does not modify previously imported orders so modifications are safe from
being re-written with updated import.

Some other functions are usefull to manage delivery:

 * create labels for shipping
 * export for delivery start-up Cubyn and Wing
 * plot delivery addresses on Google maps
 * generate invoices
 * parse address with Google API
 
It is an development version and is not supposed to be optimised, in short _it just works_.
