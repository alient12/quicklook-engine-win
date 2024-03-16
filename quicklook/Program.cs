using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Shell32;
using WebSocketSharp;
using WebSocketSharp.Server;
using CommandLine;
using CommandLine.Text;

namespace quicklook
{
    internal class Program
    {

        static void Main(string[] args)
        {
            Console.OutputEncoding = System.Text.Encoding.UTF8;

            var parser = new CommandLine.Parser(with => with.HelpWriter = Console.Out);
            var parserResult = parser.ParseArguments<Options>(args);

            parserResult.WithParsed<Options>(o =>
            {
                var wssv = new WebSocketServer(o.Port);
                wssv.AddWebSocketService<Server>("/websocket");
                wssv.Start();

                if (wssv.IsListening)
                {
                    Console.WriteLine("Listening on port {0}, and providing WebSocket services:", wssv.Port);

                    foreach (var path in wssv.WebSocketServices.Paths)
                        Console.WriteLine("- {0}", path);
                }

                Globals.server = wssv;
                InterceptKeys.Start();
            });
        }

    }

    public class Server : WebSocketBehavior
    {
        protected override void OnOpen()
        {
            base.OnOpen();
            Globals.sessions.Add(ID, this);
        }

        protected override void OnClose(CloseEventArgs e)
        {
            base.OnClose(e);
            Globals.sessions.Remove(ID);
        }

        protected override void OnMessage(MessageEventArgs e)
        {
            if (e.Data == "terminate")
            {
                Console.WriteLine("Stopping server...");
                Globals.server.Stop();
                Environment.Exit(0);
            }
            else if (e.Data == "close")
            {
                Globals.isClicked = false;
                InterceptKeys.timer.Stop();
            }
        }

        public void SendMessage(string message)
        {
            Send(message);
        }
    }

    public class Options
    {
        [Option('p', "port", Required = false, HelpText = "Set port.")]
        public int Port { get; set; } = 6969;
    }
}
