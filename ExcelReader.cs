using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

public class ExcelReader : IDisposable
{
	private SpreadsheetDocument doc;
	private OpenXmlReader reader;
	private Dictionary<string, string> sharedStrings = new Dictionary<string, string>();
	private IList<SheetInfo> sheetInfos = new List<SheetInfo>();

	public IList<SheetInfo> Sheets
	{
		get { return sheetInfos; }
	}

	public ExcelReader( string path )
	{
		doc = SpreadsheetDocument.Open( path, false );
		Initialize();
	}

	public ExcelReader( Stream stream )
	{
		doc = SpreadsheetDocument.Open( stream, false );
		Initialize();
	}

	private void Initialize()
	{
		var wbPart = doc.WorkbookPart;
		var ssPart = wbPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();

		if ( ssPart != null )
		{
			using ( var ssReader = OpenXmlReader.Create( ssPart ) )
			{
				var idx = 0;

				while ( ssReader.Read() )
				{
					if ( ssReader.ElementType != typeof( Text ) || ssReader.IsEndElement )
						continue;

					sharedStrings[idx.ToString()] = ssReader.GetText();
					idx++;
				}
			}
		}

		using ( var wbReader = OpenXmlReader.Create( wbPart ) )
		{
			var sheetIndex = 0;
			while ( wbReader.Read() )
			{
				if ( wbReader.ElementType == typeof( Sheet ) && wbReader.IsStartElement )
				{
					string rid = null;
					string id = null;
					string name = null;

					foreach ( var a in wbReader.Attributes )
					{
						switch ( a.LocalName )
						{
							case "sheetId":
								id = a.Value;
								break;
							case "name":
								name = a.Value;
								break;
							case "id":
								rid = a.Value;
								break;
						}
					}

					if ( new[] { name, id, rid }.Any( string.IsNullOrEmpty ) == false )
					{
						sheetInfos.Add( new SheetInfo
							{
								Index = sheetIndex,
								Name = name,
								Id = id,
								RId = rid
							} );
					}

					sheetIndex++;
				}
			}
		}
	}

	public void Dispose()
	{
		if ( reader != null )
			reader.Dispose();
		doc.Dispose();
	}

	public IEnumerable<CellDataCollection> ReadSheet( int sheetIndex )
	{
		var sheetInfo = sheetInfos.FirstOrDefault( x => x.Index == sheetIndex );
		if ( sheetInfo == null )
		{
			throw new ArgumentOutOfRangeException( "sheetIndex" );
		}

		return ReadSheet( sheetInfo );
	}

	public IEnumerable<CellDataCollection> ReadSheet( string sheetName )
	{
		var sheetInfo = Sheets.FirstOrDefault( x => x.Name.Equals( sheetName, StringComparison.InvariantCultureIgnoreCase ) );
		if ( sheetInfo == null )
		{
			throw new ArgumentException(
				string.Format( "File does not contain a sheet named {0}", sheetName ),
				"sheetName" );
		}

		return ReadSheet( sheetInfo );
	}

	public IEnumerable<CellDataCollection> ReadSheet( SheetInfo sheetInfo )
	{
		if ( sheetInfo == null )
		{
			throw new ArgumentNullException( "sheetInfo" );
		}

		var wsPart = doc.WorkbookPart.GetPartById( sheetInfo.RId );

		reader = OpenXmlReader.Create( wsPart, false );

		Func<string, string> getAttr = n =>
		{
			if ( reader.Attributes.Count == 0 )
				return null;

			var a = reader.Attributes.FirstOrDefault( x => x.LocalName == n );
			return a.Value;
		};

		Func<Type, Type, bool> readToNext = ( t, pt ) =>
		{
			while ( reader.Read() )
			{
				if ( reader.ElementType == t && reader.IsStartElement || reader.ElementType == pt && reader.IsEndElement )
					break;
			}

			return reader.EOF ? false : reader.ElementType == t && reader.IsStartElement;
		};

		while ( readToNext( typeof( Row ), typeof( SheetData ) ) )
		{
			var cellValues = new CellDataCollection();

			while ( readToNext( typeof( Cell ), typeof( Row ) ) )
			{
				var cellReference = getAttr( "r" );
				var isSharedString = getAttr( "t" ) == "s";

				//TODO: implement support for number formats and date detection
				// var sAttr = getAttr( "s" );

				readToNext( typeof( CellValue ), typeof( Cell ) );

				CellData obj;
				var data = reader.GetText();
				if ( isSharedString )
				{
					obj = new CellData { Value = sharedStrings[data], ValueType = typeof( string ) };
				}
				else if ( string.IsNullOrWhiteSpace( data ) )
				{
					obj = new CellData { Value = data, ValueType = typeof( string ) };
				}
				else
				{
					obj = new CellData { Value = data };
					obj.ValueType = typeof( decimal );
				}

				obj.Coordinates = cellReference;

				cellValues.Add( obj );
			}

			yield return cellValues;
		}
	}

	public class CellDataCollection : List<CellData>
	{
		public CellDataCollection()
		{
		}

		public CellDataCollection( int capacity )
			: base( capacity )
		{
		}

		public CellDataCollection( IEnumerable<CellData> collection )
			: base( collection )
		{
		}

		public CellData Get( string columnName )
		{
			return this.FirstOrDefault( x => x.Column.Equals( columnName, StringComparison.InvariantCultureIgnoreCase ) );
		}
	}

	public class CellData
	{
		private string coordinates;

		public string Coordinates
		{
			get { return coordinates; }
			internal set
			{
				coordinates = value;
				var parts = Regex.Split( coordinates, @"(\d+)" );
				Row = Int32.Parse( parts[1] );
				Column = parts[0];
			}
		}

		public int Row { get; private set; }
		public string Column { get; private set; }

		public Type ValueType { get; internal set; }
		public string Value { get; internal set; }

		public bool IsString
		{
			get { return ValueType == typeof( string ); }
		}

		public bool IsDecimal
		{
			get { return ValueType == typeof( decimal ); }
		}

		public override string ToString()
		{
			return Value;
		}

		public decimal ToDecimal()
		{
			return Convert.ToDecimal( Value, CultureInfo.InvariantCulture );
		}

		public decimal ToInt()
		{
			return Convert.ToInt32( Value, CultureInfo.InvariantCulture );
		}

		public DateTime ToDateTime()
		{
			return DateTime.FromOADate( Convert.ToDouble( Value, CultureInfo.InvariantCulture ) );
		}

		public static implicit operator string( CellData t )
		{
			return t == null ? null : t.Value;
		}
	}

	public class SheetInfo
	{
		public int Index { get; set; }
		public string Name { get; set; }
		internal string Id { get; set; }
		internal string RId { get; set; }
	}
}