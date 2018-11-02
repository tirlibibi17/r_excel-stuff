# An vendor-category map that you can fill in manually (incremental load)

A solution to [Power Pivot - Rollup new data into categories upon each refresh : excel](https://www.reddit.com/r/excel/comments/9tdwkv/power_pivot_rollup_new_data_into_categories_upon/).

Download the file [here](https://github.com/tirlibibi17/r_excel-stuff/raw/master/9tdwkv/Incremental%20Vendor%20Map.xlsx).

[This video](https://github.com/tirlibibi17/r_excel-stuff/raw/master/9tdwkv/demo.mp4) shows how it works:

* the transaction table is updated with a purchase from a new vendor
* refreshing all shows a blank category in the enriched transaction table
* the vendor-category table has a new entry with a blank category
* we update the category manually and refresh all again
* the enriched transaction table now has the new category and the vendor-category mapping table has kept the change we made manually