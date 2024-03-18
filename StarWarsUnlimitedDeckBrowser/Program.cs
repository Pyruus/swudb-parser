using StarWarsUnlimitedDeckBrowser.CollectionParser;

var collection = Parser.ParseCollection(@"--XLSX FILE LOCATION--");
Parser.ConvertCollectionToCsv(@"--PROJECT LOCATION--\Results\collection.csv", collection);