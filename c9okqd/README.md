# Using the openweathermap API with Power Query

A solution to [Having trouble with an API](https://www.reddit.com/r/excel/comments/c9okqd/having_trouble_with_an_api/).

Watch the video [here](https://youtu.be/CGEnvGP3XLI).

Start off with my [PQ Template](https://github.com/tirlibibi17/excel-pq/tree/master/PQ%20Template) and paste this code into a blank query:

	let
		Source = Json.Document(Web.Contents("https://api.openweathermap.org/data/2.5/weather?zip=94104,us&appid=<KEY>")),
		sunset = DateTime.From(Source[sys][sunset] / 86400 + 25569 + Source[timezone]/3600/24),
		weather = Source[weather]{0}[description],
		#"Converted to Table" = #table({"weather","sunset local time"}, {{weather,sunset}}),
		#"Changed Type" = Table.TransformColumnTypes(#"Converted to Table",{{"weather", type text}, {"sunset local time", type datetime}})
	in
		#"Changed Type"
		
(you'll need to get API key first)