// Define __construct, which will be used for class inheritance interop (es5 class extends es6 class).
// see https://github.com/Microsoft/TypeScript/issues/15397
var __construct =
    (this && this.__construct) ||
    (typeof Reflect !== "undefined" && Reflect.construct
        ? function(s, t, a) {
              return t !== null ? Reflect.construct(t, a, s.constructor) : s;
          }
        : function(s, t, a) {
              return (t !== null && t.apply(s, a)) || s;
          });

VSS.init({ explicitNotifyLoaded: true, usePlatformScripts: true, usePlatformStyles: true, extensionReusedCallback: registerContribution });

// We need to register the new contribution if this extension host is reused
function registerContribution(contribution) {
    // Register the fully-qualified contribution id here.
    // Because we're using the contribution id, we do NOT need to define a registeredObjectId in the extension manfiest.
    VSS.register("sample-wiki-menu", () => {
        return {
            execute: async context => {
                var item = context.treeItem;
                var wiki = context.wiki;
                VSS.require(["VSS/Service", "TFS/Wiki/WikiRestClient"], function(VSS_Service, TFS_Wiki_WebApi) {
                    // Get the REST client
                    var wikiClient = VSS_Service.getClient(TFS_Wiki_WebApi.WikiHttpClient);
                    wikiClient.getPageText(wiki.projectId, wiki.id, item.fullName).then(function(pageContent) {
                        VSS.getService("ms.vss-wiki-web.wiki-extension-data-service").then(dataService => {
                            dataService.getHtmlContent(pageContent).then(htmlContent => {
                                var pdf = new jsPDF("p", "pt", "a4");
                                var pageWidth = pdf.internal.pageSize.getWidth();
                                var margin = 25;
                                pdf.setFontSize(18);
                                pdf.fromHTML(
                                    htmlContent,
                                    margin, // x coord
                                    margin,
                                    {
                                        // y coord
                                        width: pageWidth - 2 * margin // max width of content on PDF
                                    }
                                );
                                pdf.save(item.name + ".pdf");
                            });
                            return dataService;
                        });
                        /* 
                        console.log(pageContent);
                        var doc = new jsPDF({ orientation: "portrait", unit: "mm", lineHeight: 1 });
                        var margin = 10;
                        var textHeight = 4;
                        var pageHeight = doc.internal.pageSize.getHeight();
                        var pageWidth = doc.internal.pageSize.getWidth();
                        var splitContent = doc.splitTextToSize(pageContent, pageWidth - 2 * margin);

                        doc.setFontSize(textHeight * 4);
                        var positionY = margin;
                        for (var i = 0; i < splitContent.length; i++) {
                            doc.setFontSize(textHeight * 4);
                            if (positionY + textHeight > pageHeight - margin) {
                                doc.addPage();
                                positionY = margin;
                            }
                            doc.text(splitContent[i], margin, positionY + textHeight);
                            positionY += textHeight + 4;
                        }
                        doc.save(item.name + ".pdf"); */
                    });
                    // ...
                });
            }
        };
    });
}

// Show context info when ready
VSS.ready(function() {
    registerContribution(VSS.getContribution());
    VSS.notifyLoadSucceeded();
});
