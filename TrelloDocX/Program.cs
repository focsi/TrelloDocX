using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Manatee.Trello;
using Manatee.Trello.ManateeJson;
using Manatee.Trello.WebApi;
using Novacode;

namespace TrelloDocX
{
    class Program
    {
        static void Main( string[] args )
        {
            Console.WriteLine( "Start" );
            try
            {
                if( args.Count() < 3 )
                {
                    Console.WriteLine( "Missing parameter(s). Usage: TrelloDocX UserToken BoardId Output.docx" );
                    Console.WriteLine( $"Get user token: {Constants.AuthorizeURL}{Constants.AppKey}" );
                    return;
                }

                InitConfig( args[0] );

                var board = new Board( args[1] );
                WriteDocX( board, args[2] );
            }

            catch( Exception ex )
            {
                Console.WriteLine( $"Error: {ex.Message} " );
            }
            Console.WriteLine( "End" );
        }

        private static void InitConfig( string userToken )
        {
            var serializer = new ManateeSerializer();
            TrelloConfiguration.Serializer = serializer;
            TrelloConfiguration.Deserializer = serializer;
            TrelloConfiguration.JsonFactory = new ManateeFactory();
            TrelloConfiguration.RestClientProvider = new WebApiClientProvider();
            TrelloAuthorization.Default.AppKey = Constants.AppKey;
            TrelloAuthorization.Default.UserToken = userToken;
        }

        private static void WriteDocX( Board board, string docPath )
        {
            try
            {
                // Create a document in memory:
                var doc = DocX.Create( docPath );

                foreach( var list in board.Lists )
                {
                    Paragraph paragraph = doc.InsertParagraph( list.Name );
                    paragraph.StyleName = "Heading1";

                    var cards = list.Cards;
                    foreach( var card in cards )
                    {
                        paragraph = doc.InsertParagraph( card.Name );
                        paragraph.StyleName = "Heading5";

                        var comments = card.Comments.Select( c => c.Data.Text ).ToArray();
                        InsertBulletedList( doc, comments );
                    }
                }

                doc.Save();
            }
            catch( Exception ex )
            {
                Console.WriteLine( $"DocX saving error: {ex.Message} " );
            }

            //            Process.Start( "WINWORD.EXE", docPath );
        }

        private static void InsertBulletedList( DocX doc, string[] comments )
        {
            if( comments.Count() > 0 )
            {
                var bulletedList = doc.AddList( comments[0], 0, ListItemType.Bulleted );
                for( int i = 1; i < comments.Count(); i++ )
                    doc.AddListItem( bulletedList, comments[i] );

                doc.InsertList( bulletedList );
            }
        }
    }
}
