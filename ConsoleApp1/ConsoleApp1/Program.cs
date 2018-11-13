using System;
using System.Threading.Tasks;
using Discord;
using Discord.Commands;
using Discord.WebSocket;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Sheets.v4;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using System.IO;
using System.Threading;
using Microsoft.Extensions.DependencyInjection;
using System.Reflection;
using Google.Apis.Sheets.v4.Data;
using Data = Google.Apis.Sheets.v4.Data;
using System.Collections.Generic;
using System.Linq;

namespace MyBot
{
    public class Program
    {
        // CommandService and Service Collection for use with commands
        private IServiceCollection _map = new ServiceCollection();
        private CommandService _commands = new CommandService();
        public static DiscordSocketClient _client;
        public static CommandContext context;
        private IServiceProvider services;

        //SpreadSheet Stuff


        // Program Entry Point
        public static void Main(string[] args) => new Program().MainAsync().GetAwaiter().GetResult();


        private static Task Logger(LogMessage message)
        {
            var cc = Console.ForegroundColor;
            switch (message.Severity)
            {
                case LogSeverity.Critical:
                case LogSeverity.Error:
                    Console.ForegroundColor = ConsoleColor.Red;
                    break;
                case LogSeverity.Warning:
                    Console.ForegroundColor = ConsoleColor.Yellow;
                    break;
                case LogSeverity.Info:
                    Console.ForegroundColor = ConsoleColor.White;
                    break;
                case LogSeverity.Verbose:
                case LogSeverity.Debug:
                    Console.ForegroundColor = ConsoleColor.DarkGray;
                    break;
            }
            Console.WriteLine($"{DateTime.Now,-19} [{message.Severity,8}] {message.Source}: {message.Message}");
            Console.ForegroundColor = cc;
            return Task.CompletedTask;
        }

        public async Task MainAsync()
        {
            // Discord Auth Stuff
            string token = "NTEwMTk2NjcxMTAwMjIzNTA1.DsY1iw.SpU-9T6FkXNHdB96ZaUVBKSoos8";
            _client = new DiscordSocketClient();
            services = new ServiceCollection().BuildServiceProvider();
            _client.Log += Logger;
            await InstallCommands();
            //Google Auth Stuff
            AuthorizeGoogle();
            Commands.sheetid = "1GbhI5Gla_BrCMK8fZIgs_eTbSDicPRTk77x-cfzeGAw";
            Commands.range = "A1";
            Commands.valueinputoption = (SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.USERENTERED);
            Commands.insertdataoption = (SpreadsheetsResource.ValuesResource.AppendRequest.InsertDataOptionEnum.INSERTROWS);
            await _client.LoginAsync(TokenType.Bot, token);
            await _client.StartAsync();
            await Task.Delay(-1);
        }
        private static SheetsService AuthorizeGoogle()
        {
            using (var stream = new FileStream("credentials.json", FileMode.Open, FileAccess.Read))
            {
                string credpath = System.Environment.GetFolderPath(System.Environment.SpecialFolder.Personal);
                credpath = Path.Combine(credpath, ".credentials/sheets.googleapis.com-dotnet-quickstart.json");
                Commands.credential = GoogleWebAuthorizationBroker.AuthorizeAsync(GoogleClientSecrets.Load(stream).Secrets, Commands.Scopes, "user", CancellationToken.None, new FileDataStore(credpath, true)).Result;
                Console.WriteLine("Authozied! Credential filed saved to: " + credpath);
            }
            Commands.service = new SheetsService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = Commands.credential,
                ApplicationName = Commands.ApplicationName,
            });
            return Commands.service;
        }
        public async Task InstallCommands()
        {
            _client.MessageReceived += HandleCommandAsync;
            await _commands.AddModulesAsync(Assembly.GetEntryAssembly());
        }

        private async Task HandleCommandAsync(SocketMessage arg)
        {
            var message = arg as SocketUserMessage;
            if (message == null) return;
            // Create a number to track where the prefix ends and the command begins
            int argPos = 0;
            // Determine if the message is a command, based on if it starts with '!' or a mention prefix
            if (!(message.HasCharPrefix('!', ref argPos) || message.HasMentionPrefix(_client.CurrentUser, ref argPos))) return;
            // Create a Command Context
            context = new CommandContext(_client, message);
            // Execute the command. (result does not indicate a return value, 
            // rather an object stating if the command executed successfully)
            var result = await _commands.ExecuteAsync(context, argPos, services);
            if (!result.IsSuccess)
                await context.Channel.SendMessageAsync(result.ErrorReason);
        }
    }

    public class Commands : ModuleBase
    {
        public static SheetsService service;
        public static Data.ValueRange requestbody;
        public static UserCredential credential;
        public static string sheetid, range;
        public static string[] Scopes = { SheetsService.Scope.Spreadsheets };
        public static string ApplicationName = "DevArea Client";
        public static SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum valueinputoption;
        public static SpreadsheetsResource.ValuesResource.AppendRequest.InsertDataOptionEnum insertdataoption;

        private static int GetRowByDiscord(string discordhandle)
        {
            range = "B2:B";
            SpreadsheetsResource.ValuesResource.GetRequest get = service.Spreadsheets.Values.Get(sheetid, range);
            ValueRange response = get.Execute();
            if (response.Values == null) { return 0; }
            IList<IList<object>> values = response.Values;
            int rownum = 2;
            foreach (var row in values)
            {
                foreach (var cell in row)
                {
                    if ((string)cell == discordhandle)
                    {
                        return rownum;
                    }
                    rownum++;
                }
            }
            return 0;
        }
        private static int GetRowByName(string name)
        {
            range = "A2:A";
            SpreadsheetsResource.ValuesResource.GetRequest get = service.Spreadsheets.Values.Get(sheetid, range);
            ValueRange response = get.Execute();
            if (response.Values == null) { return 0; }
            IList<IList<object>> values = response.Values;
            int rownum = 2;
            foreach (var row in values)
            {
                foreach (var cell in row)
                {
                    if ((string)cell == name)
                    {
                        return rownum;
                    }
                    rownum++;
                }
            }
            return 0;
        }
        // ~say hello -> hello
        [Command("say"), Summary("Echos a message.")]
        public async Task Say([Remainder, Summary("The text to echo")] string echo)
        {
            // ReplyAsync is a method on ModuleBase
            await ReplyAsync(echo);
        }

        [Command("halp"), Summary("Help list for commands.")]
        public async Task Halp()
        {
            var builder = new EmbedBuilder();
            builder.WithTitle("Command Information");
            builder.WithDescription("Outline of commands and their parameters.");
            builder.AddInlineField("!halp", "This menu!");
            builder.AddInlineField("!sheet", "Displays the google sheet url.");
            builder.AddInlineField("!add <character>", "Add your main character to the spreadsheet.");
            builder.AddInlineField("!role <Tank/DPS/Healer>", "Sets your raid role");
            builder.AddInlineField("!class <ingame class>", "Set your in-game class.");
            builder.AddInlineField("!spec <ingame spec>", "Set your in-game specalization.");
            builder.AddInlineField("!ilvl <###>", "Update your current item level.");
            builder.AddField("!info <character/discord>", "Displays all information about a character. Can be searched by ingame character name or Discord mention.");
            builder.WithColor(Discord.Color.Red);
            await ReplyAsync("", false, builder);
        }
        [Command("sheet"), Summary("Prints out the google sheet URL")]
        public async Task Sheet()
        {
            await ReplyAsync("Google Sheet URL: https://docs.google.com/spreadsheets/d/1GbhI5Gla_BrCMK8fZIgs_eTbSDicPRTk77x-cfzeGAw");
        }
        [Command("add"), Summary("Add yourself to the spreadsheet.")]
        public async Task Add(string name)
        {
            if (GetRowByDiscord(Program.context.User.ToString()) != 0) { await ReplyAsync("You have already added yourself."); return; }
            IList<IList<object>> data = new List<IList<object>>() { new List<object> { name, Program.context.User.ToString() } };
            range = "A1";
            requestbody = new Data.ValueRange();
            requestbody.Values = data;
            SpreadsheetsResource.ValuesResource.AppendRequest r = service.Spreadsheets.Values.Append(requestbody, sheetid, range);
            SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum valueinputoption = (SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.USERENTERED);
            SpreadsheetsResource.ValuesResource.AppendRequest.InsertDataOptionEnum insertdataoption = (SpreadsheetsResource.ValuesResource.AppendRequest.InsertDataOptionEnum.INSERTROWS);
            r.ValueInputOption = valueinputoption;
            r.InsertDataOption = insertdataoption;
            r.AccessToken = credential.Token.AccessToken;
            Data.AppendValuesResponse response = await r.ExecuteAsync();
            await ReplyAsync("Successfully added!");
        }
        [Command("role"), Summary("Set raid role.")]
        public async Task Role(string data)
        {
            if (GetRowByDiscord(Program.context.User.ToString()) == 0) { await ReplyAsync("No data found..."); return; }
            range = "C" + GetRowByDiscord(Program.context.User.ToString());
            requestbody = new Data.ValueRange();
            requestbody.MajorDimension = "ROWS";
            var info = new List<object>() { data };
            requestbody.Values = new List<IList<object>> { info };
            SpreadsheetsResource.ValuesResource.UpdateRequest update = service.Spreadsheets.Values.Update(requestbody, sheetid, range);
            update.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.RAW;
            UpdateValuesResponse result2 = update.Execute();
            await ReplyAsync("Set your role to **" + data + "**.");
        }
        [Command("class"), Summary("Set your ingame class")]
        public async Task Class(string data)
        {
            if (GetRowByDiscord(Program.context.User.ToString()) == 0) { await ReplyAsync("No data found..."); return; }
            range = "E" + GetRowByDiscord(Program.context.User.ToString());
            requestbody = new Data.ValueRange();
            requestbody.MajorDimension = "ROWS";
            var info = new List<object>() { data };
            requestbody.Values = new List<IList<object>> { info };
            SpreadsheetsResource.ValuesResource.UpdateRequest update = service.Spreadsheets.Values.Update(requestbody, sheetid, range);
            update.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.RAW;
            UpdateValuesResponse result2 = update.Execute();
            await ReplyAsync("Set your class to **" + data + "**.");

        }
        [Command("spec"), Summary("Set your main spec")]
        public async Task Spec(string data)
        {
            if (GetRowByDiscord(Program.context.User.ToString()) == 0) { await ReplyAsync("No data found..."); return; }
            range = "F" + GetRowByDiscord(Program.context.User.ToString());
            requestbody = new Data.ValueRange();
            requestbody.MajorDimension = "ROWS";
            var info = new List<object>() { data };
            requestbody.Values = new List<IList<object>> { info };
            SpreadsheetsResource.ValuesResource.UpdateRequest update = service.Spreadsheets.Values.Update(requestbody, sheetid, range);
            update.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.RAW;
            UpdateValuesResponse result2 = update.Execute();
            await ReplyAsync("Set your spec to **" + data + "**.");
        }
        [Command("ilvl"), Summary("Set your current iLevel")]
        public async Task iLvl(string data)
        {
            if (GetRowByDiscord(Program.context.User.ToString()) == 0) { await ReplyAsync("No data found..."); return; }
            range = "D" + GetRowByDiscord(Program.context.User.ToString());
            Console.WriteLine(range);
            requestbody = new Data.ValueRange();
            requestbody.MajorDimension = "ROWS";
            var info = new List<object>() { data };
            requestbody.Values = new List<IList<object>> { info };
            SpreadsheetsResource.ValuesResource.UpdateRequest update = service.Spreadsheets.Values.Update(requestbody, sheetid, range);
            update.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.RAW;
            UpdateValuesResponse result2 = update.Execute();
            await ReplyAsync("Updated your item level to: **" + data + "**.");
        }
        [Command("info"), Summary("Get information about a specific player or mention")]
        public async Task Info(string data)
        {
            IUser user = null;
            var builder = new EmbedBuilder();
            builder.WithColor(Discord.Color.Red);
            if (data.Contains("@"))
            {
                var guilds = (await Context.Client.GetGuildsAsync(CacheMode.AllowDownload));
                var users = new List<IUser>();
                foreach (var guild in guilds)
                    users.AddRange(await guild.GetUsersAsync(CacheMode.AllowDownload));
                users = users.GroupBy(o => o.Id).Select(o => o.First()).ToList();
                var search = users.Where(o => o.Username.ToLower().Contains(data.ToLower()) || Context.Message.MentionedUserIds.Contains(o.Id) || o.ToString().ToLower().Contains(data.ToLower())).ToArray();
                if (search.Length == 0 || search.Length > 1)
                {
                    await ReplyAsync("**Error!** *Unable to find that user*.");
                    return;
                }
                user = search.First();
                builder.WithTitle(user.Username.ToString() + "'s Information Page");
                SpreadsheetsResource.ValuesResource.GetRequest get;
                ValueRange response;
                IList<IList<object>> values;
                string[] types = { "Role", "iLevel", "Class", "Spec", "Attendance" };
                string[] fields = { "N/A", "N/A", "N/A", "N/A", "N/A" };
                int i = 0;
                for (char c = 'C'; c <= 'G'; c++)
                {
                    range = c + GetRowByDiscord(user.ToString()).ToString();
                    get = service.Spreadsheets.Values.Get(sheetid, range);
                    response = get.Execute();
                    values = response.Values;
                    if (values != null && values.Count > 0)
                    {
                        foreach (var value in values)
                        {
                            if (value[0] == null) { break; }
                            Console.WriteLine(value[0]);
                            fields[i] = (string)value[0];
                        }
                    }
                    i++;
                }
                i = 0;
                foreach (string item in fields)
                {
                    builder.AddInlineField(types[i], item);
                    i++;
                }
                await ReplyAsync("", false, builder);
            }
            else
            {
                builder.WithTitle(data + "'s Information Page");
                SpreadsheetsResource.ValuesResource.GetRequest get;
                ValueRange response;
                IList<IList<object>> values;
                string[] types = { "Role", "iLevel", "Class", "Spec", "Attendance" };
                string[] fields = { "", "", "", "", "N/A" };
                int i = 0;
                for (char c = 'C'; c <= 'G'; c++)
                {
                    range = c + GetRowByName(data).ToString();
                    get = service.Spreadsheets.Values.Get(sheetid, range);
                    response = get.Execute();
                    values = response.Values;
                    if (values != null && values.Count > 0)
                    {
                        foreach (var value in values)
                        {
                            if (value[0] == null) { break; }
                            Console.WriteLine(value[0]);
                            fields[i] = (string)value[0];
                        }
                    }
                    i++;
                }
                i = 0;
                foreach(string item in fields)
                {
                    builder.AddInlineField(types[i], item);
                    i++;
                }
                await ReplyAsync("", false, builder);
            }
        }
    }
}