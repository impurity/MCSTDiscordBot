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
using MCSTDiscord;
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
        static void CheckAlive()
        {
            while (true)
            {
                Thread.Sleep(300000); //5 Minutes
                System.Diagnostics.Debug.WriteLine("Refreshed\n");
                AuthorizeGoogle();
            }
        }

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
            ThreadStart AliveThread = new ThreadStart(CheckAlive);

            Thread t = new Thread(AliveThread);
            t.Start();
            string token = "NTEwMTk2NjcxMTAwMjIzNTA1.DsY1iw.SpU-9T6FkXNHdB96ZaUVBKSoos8";
            _client = new DiscordSocketClient();
            services = new ServiceCollection().BuildServiceProvider();
            _client.Log += Logger;
            await InstallCommands();
            //Google Auth Stuff
            AuthorizeGoogle();
            Commands.sheetid = "1GbhI5Gla_BrCMK8fZIgs_eTbSDicPRTk77x-cfzeGAw";
            Commands.range = "B5";
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
            builder.WithDescription("Outline of commands and their parameters. Start by adding your character to the sheet using the !add <Player Name> command.");
            builder.AddField("!add <character>", "Add your main character to the spreadsheet.");
            builder.AddField("!alt <character> <item level>", "Add up to two alts and their item levels.");
            builder.AddInlineField("!halp", "This menu!");
            //builder.AddInlineField("!sheet", "Displays the google sheet url.");
            builder.AddInlineField("!raid <Tank | Healer | RDPS | MDPS | DPS>", "Sets your raid role");
            builder.AddInlineField("!class <Warrior | Mage | Priest | Hunter | Warlock | Demon Hunter | Death Knight | Paladin | Druid | Rogue | Shaman>", "Set your in-game class.");
            builder.AddInlineField("!spec <ingame spec>", "Set your in-game specalization.");
            builder.AddInlineField("!ilvl <###>", "Update your current item level.");
            builder.AddField("!lookup <character/discord>", "Displays all information about a character. Can be searched by ingame character name or Discord mention.");
            builder.AddField("!list", "Lists all the names and discords that are available to lookup with the !lookup command.");
            builder.WithColor(Discord.Color.Red);
            await ReplyAsync("", false, builder);
        }
        //[Command("sheet"), Summary("Prints out the google sheet URL")]
        //public async Task Sheet()
        //{
        //    await ReplyAsync("Google Sheet URL: https://docs.google.com/spreadsheets/d/1GbhI5Gla_BrCMK8fZIgs_eTbSDicPRTk77x-cfzeGAw");
        //}
        [Command("add"), Summary("Add yourself to the spreadsheet.")]
        public async Task Add(string name)
        {
            if (Utils.GetRowByDiscord(Program.context.User.ToString()) != 0) { await ReplyAsync("You have already added yourself."); return; }
            IList<IList<object>> data = new List<IList<object>>() { new List<object> { Utils.UpperCaseIt(name.ToLower()), Program.context.User.ToString() } };
            range = "B5";
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
        [Command("alt"), Summary("Add an alt to the spreadsheet with its item level.")]
        public async Task Alt(string name, string ilvl)
        {
            range = "H" + Utils.GetRowByDiscord(Program.context.User.ToString());
            if (Utils.GetRowByDiscord(Program.context.User.ToString()) == 0) { await ReplyAsync("No user found, be sure you've added a main character with the !add *<Player Name>* command"); return; }
            if (Utils.GetRowByAlt(1, name) != 0 || Utils.GetRowByAlt(2, name) != 0) { await ReplyAsync("You've already added this alt."); return; }
            if (ilvl.All(char.IsDigit) == false || ilvl.Length > 3) { await ReplyAsync("Invalid item level."); return; }
            if (Utils.CellEmpty("H" + Utils.GetRowByDiscord(Program.context.User.ToString())) == false && Utils.CellEmpty("J" + Utils.GetRowByDiscord(Program.context.User.ToString())) == false) { await ReplyAsync("You can only have up to two alt's, if you need an alt removed contact an Officer."); return; }
            if (Utils.GetRowByAlt(1, name) == 0)
            {
                if (Utils.CellEmpty(range) == true)
                {
                    range = "H" + Utils.GetRowByDiscord(Program.context.User.ToString());
                }
                else
                {
                    if (Utils.GetRowByAlt(2, name) == 0)
                    {
                        range = "J" + Utils.GetRowByDiscord(Program.context.User.ToString());
                    }
                }
            }
            IList<IList<object>> data = new List<IList<object>>() { new List<object> { Utils.UpperCaseIt(name.ToLower()), ilvl } };
            requestbody = new Data.ValueRange();
            requestbody.MajorDimension = "ROWS";
            requestbody.Values = data;
            SpreadsheetsResource.ValuesResource.UpdateRequest update = service.Spreadsheets.Values.Update(requestbody, sheetid, range);
            update.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.RAW;
            UpdateValuesResponse result = update.Execute();
            await ReplyAsync("Successfully added **" + Utils.UpperCaseIt(name.ToLower()) + "** as an alt.");
        }
        [Command("raid"), Summary("Set raid role.")]
        public async Task Raid(string data)
        {
            bool valid = false;
            List<string> roles = new List<string>() { "Tank", "Healer", "RDPS", "MDPS", "DPS" };
            foreach (string role in roles)
            {
                if (data.ToLower() == role.ToLower()) { valid = true; data = role; }
            }
            if (valid == false) { await ReplyAsync("Invalid parameter: **" + data + "**. Please type !halp to see a list of supported parameters."); return; }
            if (Utils.GetRowByDiscord(Program.context.User.ToString()) == 0) { await ReplyAsync("No user found, be sure you've added yourself with the !add *<Player Name>* command"); return; }
            range = "D" + Utils.GetRowByDiscord(Program.context.User.ToString());
            requestbody = new Data.ValueRange();
            requestbody.MajorDimension = "ROWS";
            var info = new List<object>() { data };
            requestbody.Values = new List<IList<object>> { info };
            SpreadsheetsResource.ValuesResource.UpdateRequest update = service.Spreadsheets.Values.Update(requestbody, sheetid, range);
            update.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.RAW;
            UpdateValuesResponse result2 = update.Execute();
            await ReplyAsync("Your raid role has been set to **" + data + "**.");
        }
        [Command("class"), Summary("Set your ingame class")]
        public async Task Class([Remainder, Summary("Your ingame class.")] string data)
        {
            List<string> gameclasses = new List<string>() { "Warrior", "Mage", "Priest", "Hunter", "Warlock", "Monk", "Demon Hunter", "Death Knight", "Paladin", "Druid", "Rogue", "Shaman" };
            bool valid = false;
            foreach (string classes in gameclasses)
            {
                if (data.ToLower() == classes.ToLower()) { valid = true; data = classes; }
            }
            if (valid == false) { await ReplyAsync("Invalid parameter: **" + data + "**. Please type !halp to see a list of supported parameters."); return; }
            if (Utils.GetRowByDiscord(Program.context.User.ToString()) == 0) { await ReplyAsync("No user found, be sure you've added yourself with the !add *<Player Name>* command"); return; }
            range = "F" + Utils.GetRowByDiscord(Program.context.User.ToString());
            requestbody = new Data.ValueRange();
            requestbody.MajorDimension = "ROWS";
            var info = new List<object>() { data };
            requestbody.Values = new List<IList<object>> { info };
            SpreadsheetsResource.ValuesResource.UpdateRequest update = service.Spreadsheets.Values.Update(requestbody, sheetid, range);
            update.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.RAW;
            UpdateValuesResponse result2 = update.Execute();
            await ReplyAsync("Your class has been set to **" + data + "**.");

        }
        [Command("spec"), Summary("Set your main spec")]
        public async Task Spec([Remainder, Summary("Set your main specialization ingame.")] string data)
        {
            bool valid = false;
            List<string> specs = new List<string>() { "Arms", "Fury", "Prot", "Protection", "Arcane", "Fire", "Frost", "Discipline", "Holy", "Shadow", "Beast Mastery", "Marskman", "Survival", "Affliction", "Demonology", "Destruction", "Brewmaster", "Mistweaver", "Windwalker", "Havoc", "Vengeance", "Blood", "Unholy", "Retribution", "Balance", "Feral", "Guardian", "Restoration", "Assassination", "Outlaw", "Subtlety", "Elemental", "Enhancement" };
            foreach (string spec in specs)
            {
                if (data.ToLower() == "prot") { valid = true; data = "Protection"; return; }
                if (data.ToLower() == spec.ToLower()) { valid = true; data = spec; }
            }
            if (valid == false) { await ReplyAsync("Invalid parameter: **" + data + "**. Please type !halp to see a list of supported parameters."); return; }
            if (Utils.GetRowByDiscord(Program.context.User.ToString()) == 0) { await ReplyAsync("No user found, be sure you've added yourself with the !add *<Player Name>* command"); return; }
            range = "G" + Utils.GetRowByDiscord(Program.context.User.ToString());
            requestbody = new Data.ValueRange();
            requestbody.MajorDimension = "ROWS";
            var info = new List<object>() { data };
            requestbody.Values = new List<IList<object>> { info };
            SpreadsheetsResource.ValuesResource.UpdateRequest update = service.Spreadsheets.Values.Update(requestbody, sheetid, range);
            update.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.RAW;
            UpdateValuesResponse result2 = update.Execute();
            await ReplyAsync("Your spec has been set to **" + data + "**.");
        }
        [Command("ilvl"), Summary("Set your current iLevel")]
        public async Task iLvl(string data)
        {
            bool valid = true;
            if (data.Length > 3 || !data.All(char.IsDigit)) { valid = false; }
            if (valid == false) { await ReplyAsync("Invalid parameter: **" + data + "**. Please type !halp to see a list of supported parameters."); return; }
            if (Utils.GetRowByDiscord(Program.context.User.ToString()) == 0) { await ReplyAsync("No user found, be sure you've added yourself with the !add *<Player Name>* command"); return; }
            range = "E" + Utils.GetRowByDiscord(Program.context.User.ToString());
            Console.WriteLine(range);
            requestbody = new Data.ValueRange();
            requestbody.MajorDimension = "ROWS";
            var info = new List<object>() { data };
            requestbody.Values = new List<IList<object>> { info };
            SpreadsheetsResource.ValuesResource.UpdateRequest update = service.Spreadsheets.Values.Update(requestbody, sheetid, range);
            update.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.RAW;
            UpdateValuesResponse result2 = update.Execute();
            await ReplyAsync("Your item level has been set to: **" + data + "**.");
        }
        [Command("lookup"), Summary("Get information about a specific player or mention")]
        public async Task Lookup(string data)
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
                if (Utils.GetRowByDiscord(user.ToString()) == 0) { await ReplyAsync("No user found, be sure they have been added with the !add command before querying"); return; };
                builder.WithTitle(user.Username.ToString() + "'s Information Page");
                SpreadsheetsResource.ValuesResource.GetRequest get;
                ValueRange response;
                IList<IList<object>> values;
                string[] types = { "Role", "iLevel", "Class", "Spec", "Alt Character", "Alt Character iLevel", "Alt Character", "Alt Character iLevel", "Attendance" };
                string[] fields = { "N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "N/A" };
                int i = 0;
                for (char c = 'D'; c <= 'L'; c++)
                {
                    range = c + Utils.GetRowByDiscord(user.ToString()).ToString();
                    get = service.Spreadsheets.Values.Get(sheetid, range);
                    response = get.Execute();
                    values = response.Values;
                    if (values != null && values.Count > 0)
                    {
                        foreach (var value in values)
                        {
                            if ((string)value[0] == "") { break; }
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
                if (Utils.GetRowByName(data) == 0) { await ReplyAsync("No user found, be sure they have been added with the !add command before querying"); return; };
                builder.WithTitle(Utils.UpperCaseIt(data.ToLower()) + "'s Information Page");
                SpreadsheetsResource.ValuesResource.GetRequest get;
                ValueRange response;
                IList<IList<object>> values;
                string[] types = { "Role", "iLevel", "Class", "Spec", "Alt Character", "Alt Character iLevel", "Alt Character", "Alt Character iLevel", "Attendance" };
                string[] fields = { "N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "N/A" };
                int i = 0;
                for (char c = 'D'; c <= 'L'; c++)
                {
                    range = c + Utils.GetRowByName(data).ToString();
                    get = service.Spreadsheets.Values.Get(sheetid, range);
                    response = get.Execute();
                    values = response.Values;
                    if (values != null && values.Count > 0)
                    {
                        foreach (var value in values)
                        {
                            if ((string)value[0] == "") { break; }
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
        }
        [Command("list"), Summary("List all players in the spreadsheet.")]
        public async Task List()
        {
            var builder = new EmbedBuilder();
            builder.WithTitle("Character List:");
            builder.WithDescription("All available characters to query with the !lookup command");
            range = "B5:C";
            SpreadsheetsResource.ValuesResource.GetRequest get = Commands.service.Spreadsheets.Values.Get(Commands.sheetid, range);
            ValueRange response = get.Execute();
            if (response.Values == null) { await ReplyAsync("**Error**: Unable to retrieve players."); return; }
            IList<IList<object>> values = response.Values;
            foreach (var row in values)
            {
                builder.AddInlineField((string)row[0], row[1]);
            }
            await ReplyAsync("", false, builder);
        }
    }
}