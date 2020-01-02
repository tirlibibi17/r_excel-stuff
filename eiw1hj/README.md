# Is there a way to feed 10^8 combinations of single digits into an algebraic formula?

A solution to [Is there a way to feed 10^8 combinations of single digits into an algebraic formula?](https://www.reddit.com/r/excel/comments/eiw1hj/is_there_a_way_to_feed_108_combinations_of_single/) posted by u/sizarieldor

Power Query on a 64-bit version of Excel to the rescue! The idea is to build one column with list `{0..9}`for each individual letter, expand each column to rows, calculate the formula in a custom column and filter that column to only keep the rows that result in 2020, and finally remove that column. Code (result loaded to tab *All combinations*):
    
    let
        Source = #table({"dummy"},{{0}}),
        #"Created Combinations" = List.Accumulate({"H","A","P","Y","N","E","W","R"}, Source, (state,current)=>let
            #"Added letter" = Table.AddColumn(state, current, each {0..9}),
            #"Expanded letter" = Table.ExpandListColumn(#"Added letter", current),
            #"Changed Type" = Table.TransformColumnTypes(#"Expanded letter",{{current, Int64.Type}})
        in
            #"Changed Type"
        ),
        #"Removed Dummy Column" = Table.RemoveColumns(#"Created Combinations",{"dummy"}),
        #"Added Custom" = Table.AddColumn(#"Removed Dummy Column", "Happy New Year", each 10000*[H] + 1000*[A] + 100*[P] + 10*[P] + [Y] +100*[N] + 10*[E] + [W] - 1000*[Y] - 100*[E] - 10*[A] - [R]),
        #"Filtered Rows" = Table.SelectRows(#"Added Custom", each [Happy New Year] = 2020),
        #"Removed Columns" = Table.RemoveColumns(#"Filtered Rows",{"Happy New Year"})
    in
        #"Removed Columns"

This took over 13 minutes to run on my 2 year old Core i5 with 8 GB of RAM.

There are 1001 solutions to the problem, 851 with H=0 and 150 with H=0. Once you know that, you can try to prove that H cannot be more that 1 before building all the combinations, to save memory and processing time. Here goes: if H is 2, you would be starting with 20000. If you maximize YEAR and make it 9999, the result is going to be more than 10001, so H=2 is impossible, and so is every other value of H higher than 1. So now, we can optimize the query to only put {0..1} in column H. Code (result loaded to tab *Optimized*)

    let
        Source = #table({"dummy"},{{0}}),
        #"Created Combinations" = List.Accumulate({"H","A","P","Y","N","E","W","R"}, Source, (state,current)=>let
            #"Added letter" = Table.AddColumn(state, current, each {0..(if current="H" then 1 else 9)}),
            #"Expanded letter" = Table.ExpandListColumn(#"Added letter", current),
            #"Changed Type" = Table.TransformColumnTypes(#"Expanded letter",{{current, Int64.Type}})
        in
            #"Changed Type"
        ),
        #"Removed Dummy Column" = Table.RemoveColumns(#"Created Combinations",{"dummy"}),
        #"Added Custom" = Table.AddColumn(#"Removed Dummy Column", "Happy New Year", each 10000*[H] + 1000*[A] + 100*[P] + 10*[P] + [Y] +100*[N] + 10*[E] + [W] - 1000*[Y] - 100*[E] - 10*[A] - [R]),
        #"Filtered Rows" = Table.SelectRows(#"Added Custom", each [Happy New Year] = 2020),
        #"Removed Columns" = Table.RemoveColumns(#"Filtered Rows",{"Happy New Year"})
    in
        #"Removed Columns"

This optimized version took a bit over 2 minutes to run, so more than 6 times faster.

Head on over [here](https://github.com/tirlibibi17/r_excel-stuff/tree/master/eiw1hj) to get the file.

I've collated a list of Power Query learning resources with the help of the community [here](/r/excel/comments/9vumd3/what_resources_would_you_recommend_for_someone/).

You will find a general introduction to Power Query on Microsoft's [support website](https://support.office.com/en-us/article/introduction-to-microsoft-power-query-for-excel-6e92e2f4-2079-4e1f-bad5-89f6269cd605).