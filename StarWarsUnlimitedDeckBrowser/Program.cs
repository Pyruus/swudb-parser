using StarWarsUnlimitedDeckBrowser.CollectionParser;

var collection = Parser.ParseCollection(@"D:\Downloads\SWU - Karty.xlsx");
Parser.ConvertCollectionToCsv(@"D:\Dev\SWUDB\StarWarsUnlimitedDeckBrowser\Results\collection.csv", collection);