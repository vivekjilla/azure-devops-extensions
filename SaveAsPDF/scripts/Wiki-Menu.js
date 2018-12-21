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
                var pdfName = item.name;
                VSS.require(["VSS/Service", "TFS/Wiki/WikiRestClient"], function(VSS_Service, TFS_Wiki_WebApi) {
                    // Get the REST client
                    var wikiClient = VSS_Service.getClient(TFS_Wiki_WebApi.WikiHttpClient);

                    wikiClient.getPage(wiki.projectId, wiki.id, item.fullName, 1).then(response => {
                        var subPagesObj = response.page.subPages;

                        var htmlContentPromises = [];
                        htmlContentPromises.push(
                            wikiClient.getPageText(wiki.projectId, wiki.id, item.fullName).then(function(pageContent) {
                                return VSS.getService("ms.vss-wiki-web.wiki-extension-data-service").then(dataService => {
                                    return dataService.getHtmlContent(pageContent).then(htmlContent => {
                                        return { title: item.fullName, content: htmlContent, index: 0 };
                                    });
                                });
                            })
                        );
                        subPagesObj.forEach((page, i) => {
                            htmlContentPromises.push(
                                wikiClient.getPageText(wiki.projectId, wiki.id, page.path).then(function(pageContent) {
                                    return VSS.getService("ms.vss-wiki-web.wiki-extension-data-service").then(dataService => {
                                        return dataService.getHtmlContent(pageContent).then(htmlContent => {
                                            return { title: page.path, content: htmlContent, index: i + 1 };
                                        });
                                    });
                                })
                            );
                        });

                        Q.all(htmlContentPromises).then(response => {
                            var dict = {};
                            response.forEach(v => (dict[v.index] = { title: v.title, content: v.content }));
                            var pdf = new jsPDF("p", "pt", "a4");
                            var pageWidth = pdf.internal.pageSize.getWidth();
                            var margin = 25;
                            pdf.setFontSize(18);

                            for (var i = 0; i < subPagesObj.length + 1; i++) {
                                pdf.text(dict[i].title, margin, margin);
                                pdf.fromHTML(
                                    dict[i].content,
                                    margin, // x coord
                                    2 * margin,
                                    {
                                        // y coord
                                        width: pageWidth - 2 * margin // max width of content on PDF
                                    }
                                );
                                if (i < subPagesObj.length) {
                                    pdf.addPage();
                                }
                            }
                            pdf.save(pdfName + ".pdf");
                        });
                    });
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
