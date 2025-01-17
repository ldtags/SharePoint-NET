using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using PnP.Core;
using PnP.Core.Auth.Services.Builder.Configuration;
using PnP.Core.Model.Security;
using PnP.Core.Model.SharePoint;
using PnP.Core.Services;
using PnP.Core.Services.Builder.Configuration;
using System.Diagnostics;

namespace SharePoint_NET {
    class Program {
        public static async Task Main(string[] args) {
            var host = Host.CreateDefaultBuilder()

            .ConfigureServices((hostingContext, services) => {
                // Add the PnP Core SDK library services
                services.AddPnPCore();

                // Add the PnP Core SDK library services configuration from the appsettings.json file
                services.Configure<PnPCoreOptions>(hostingContext.Configuration.GetSection("PnPCore"));

                // Add the PnP Core SDK Authentication Providers
                services.AddPnPCoreAuthentication();

                // Add the PnP Core SDK Authentication Providers configuration from the appsettings.json file
                services.Configure<PnPCoreAuthenticationOptions>(hostingContext.Configuration.GetSection("PnPCore"));
            })

            // Let the builder know we're running in a console
            .UseConsoleLifetime()

            // Add services to the container
            .Build();

            await host.StartAsync();
            using var scope = host.Services.CreateScope();

            // Obtain a PnP Context factory
            var pnpContextFactory = scope.ServiceProvider.GetRequiredService<IPnPContextFactory>();

            // Use the PnP Context factory to get a PnPContext for the given configuration
            using var context = await pnpContextFactory.CreateAsync("CustomInitiative");

            await AddAnonymousSharingLinks(context, "Baseline Library");

            // Cleanup console host
            host.Dispose();
        }

        public static bool HasHiddenLinkColumn(IList list) {
            foreach (var field in list.Fields) {
                if (field.InternalName == "PublicLink") {
                    return true;
                }
            }

            return false;
        }

        public static bool IsAnonymousViewLink(IGraphPermission sharingLink) {
            var link = sharingLink.Link;
            if (link.Scope != ShareScope.Anonymous) {
                return false;
            }

            if (link.Type != ShareType.View) {
                return false;
            }

            if (link.PreventsDownload) {
                return false;
            }

            return true;
        }

        public static async Task<ISharingLink?> GetAnonymousViewLink(IListItem listItem) {
            var sharingLinks = await listItem.File.GetShareLinksAsync();
            foreach (var sharingLink in sharingLinks) {
                if (IsAnonymousViewLink(sharingLink)) {
                    return sharingLink.Link;
                }
            }

            return null;
        }

        public static async Task<bool> HasAnonymousViewLink(IListItem listItem) {
            var sharingLinks = await listItem.File.GetShareLinksAsync();
            foreach (var sharingLink in sharingLinks) {
                if (IsAnonymousViewLink(sharingLink)) {
                    return true;
                }
            }

            return false;
        }

        private static async Task<ISharingLink?> AddAnonymousLink(IListItem listItem, AnonymousLinkOptions linkOptions) {
            Console.Write($"  . Creating anonymous view link for {listItem.GetDisplayName()}...");

            ISharingLink? link = null;
            ISharingLink? existingLink = await GetAnonymousViewLink(listItem);
            if (existingLink != null) {
                Console.WriteLine(" skipped, link already exists.");
                return existingLink;
            }

            try {
                link = (await listItem.CreateAnonymousSharingLinkAsync(linkOptions))?.Link;
                Console.WriteLine(" ok.");
            } catch (Exception e) {
                Console.WriteLine(" failed.");
                Console.WriteLine(e.Message);
            }

            return link;
        }

        private static async Task AddLinkToField(IListItem listItem, string internalName, ISharingLink link) {
            Console.Write($"  . Adding link to [{internalName}]...");

            IFieldUrlValue? existingLink = listItem[internalName] as FieldUrlValue;
            if (existingLink != null && existingLink.Url == link.WebUrl) {
                Console.WriteLine(" skipped, link already exists");
                return;
            }

            try {
                listItem[internalName] = new FieldUrlValue(link.WebUrl);
            } catch (KeyNotFoundException) {
                Console.WriteLine($" failed, no [{internalName}] field exists.");
                return;
            }

            try {
                await listItem.UpdateAsync();
            } catch (SharePointRestServiceException) {
                Console.WriteLine(" failed, file is likely checked out by another user");
                return;
            }

            Console.WriteLine(" ok.");
        }

        public static async Task AddAnonymousSharingLinks(IPnPContext context, string libraryName) {
            Console.WriteLine($"  . Adding anonymous sharing links to all files in {libraryName}");
            var watch = Stopwatch.StartNew();

            Console.Write($"  . Fetch list ({libraryName})...");
            var library = await context.Web.Lists.GetByTitleAsync(libraryName, p => p.RootFolder);
            if (library == null) {
                Console.WriteLine(" failed, no list found.");
                return;
            }
            Console.WriteLine(" ok.");

            Console.Write($"  . Adding [Public Link] field...");
            if (HasHiddenLinkColumn(library)) {
                Console.WriteLine(" already has a [Public Link] field.");
            } else {
                await library.Fields.AddUrlAsync("Public Link", new FieldUrlOptions() {
                    InternalName = "PublicLink",
                    AddToDefaultView = false,
                    DisplayFormat = UrlFieldFormatType.Hyperlink,
                    Description = "List column for accessing the public view link for this document",
                    Required = false
                });

                Console.WriteLine(" ok.");
            }

            var shareLinkOptions = new AnonymousLinkOptions() {
                Type = ShareType.View
            };

            await foreach (var listItem in library.Items) {
                var displayName = listItem.GetDisplayName();
                Console.Write($"  . Current file: {displayName}...");

                if (!listItem.IsFile()) {
                    Console.WriteLine(" is a folder, skipping file.");
                    continue;
                }

                Console.WriteLine(" ok.");

                ISharingLink? link = await AddAnonymousLink(listItem, shareLinkOptions);
                if (link == null) {
                    Console.WriteLine("  ! Unable to get link from link creation");
                    return;
                }

                await AddLinkToField(listItem, "PublicLink", link);
            }

            watch.Stop();
            Console.WriteLine($"  . Process complete ({watch.ElapsedMilliseconds}ms)");
        }
    }
}
