using System;
using System.ServiceModel;
using System.IO;
using Microsoft.Xrm.Sdk;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.Xrm.Sdk.Query;
using DocumentFormat.OpenXml;
using X14 = DocumentFormat.OpenXml.Office2010.Excel;
using DocumentFormat.OpenXml.Office2010.Drawing.Charts;

namespace Audit.Kpmg.Plugins.SummaryReportCreation
{
    public class ExcelDataExtractor : IPlugin
    {
        ITracingService tracing = null;
        Guid SummaryId = new Guid();
        string engagementName = "";
       public void Execute(IServiceProvider serviceProvider)
        {
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);
            tracing = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            try
            {

                tracing.Trace("Start Plugin time:" + DateTime.Now.ToString());
                //Context entity object as lead 
                //tracing.Trace("open note file");
                EntityCollection colJournalHeader = new EntityCollection();
                EntityCollection colJournalLineItem = new EntityCollection();
                bool IsAllDataValid = true;
                string ConsolidatedErrorsMain = "";
                Entity summaryData = (Entity)context.InputParameters["Target"];
                SummaryId = summaryData.Id;
                tracing.Trace("we3e Plugin time:" + summaryData.Attributes.Count);
                foreach (var ins in summaryData.Attributes)
                {
                    tracing.Trace("get 2:" + ins.Key + "----" + ins.Value);
                }
                Entity notesRecord = getJournalAttachment(summaryData.Id, service);
                tracing.Trace("get notesrecorde:" + notesRecord.Attributes.Count);
                foreach(var ins in notesRecord.Attributes)
                {
                    tracing.Trace("get 1:" + ins.Key +"----"+ins.Value);
                }


                var engagementFromReq = notesRecord.Attributes["new_engagementname"].ToString();
                var reqTypeFromReq = notesRecord.Attributes["new_requesttype"].ToString();
                string Filename = "";
                
                var fetchQuery = "<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'><entity name='new_engagement'> <link-entity name='new_confirmationrequest' from='new_engagement' to='new_engagementid' alias='new_confirmationrequest1' entityalias='undefined' link-type='outer' secondary='1'> <attribute name='new_requesttypehandingfromclient' alias='new_confirmationrequest1_new_requesttypehandingfromclient' /> <attribute name='new_confirmingparty' alias='new_confirmationrequest1_new_confirmingparty' /> <attribute name='new_amount' alias='new_confirmationrequest1_new_amount' /> <attribute name='transactioncurrencyid' alias='new_confirmationrequest1_transactioncurrencyid' /> <attribute name='new_authletterrequestdate' alias='new_confirmationrequest1_new_authletterrequestdate' /> <attribute name='new_authletterapprovaldate' alias='new_confirmationrequest1_new_authletterapprovaldate' /> <attribute name='new_brachresponserequestdate' alias='new_confirmationrequest1_new_brachresponserequestdate' /> <attribute name='new_branchresponsereceiveddate' alias='new_confirmationrequest1_new_branchresponsereceiveddate' /> <attribute name='crf69_portalamount' alias='new_confirmationrequest1_crf69_portalamount' /> <attribute name='crf69_portalcurrencies' alias='new_confirmationrequest1_crf69_portalcurrencies' /> <attribute name='crf69_currencyname' alias='new_confirmationrequest1_crf69_currencyname' /> <attribute name='new_comments' alias='new_confirmationrequest1_new_comments' /> <attribute name='crf69_confirmingpartycontact' alias='new_confirmationrequest1_crf69_confirmingpartycontact' /> <attribute name='crbcb_automaticreminderdate1' alias='new_confirmationrequest1_crbcb_automaticreminderdate1' /> <attribute name='crbcb_automaticreminderdate2' alias='new_confirmationrequest1_crbcb_automaticreminderdate2' /> <attribute name='crbcb_manualreminderdate' alias='new_confirmationrequest1_crbcb_manualreminderdate' /> <attribute name='crf69_clientcontactperson' alias='new_confirmationrequest1_crf69_clientcontactperson' /> </link-entity> <attribute name='new_name' alias='new_name' /> <filter type='and'> <condition attribute='new_engagementid' operator='eq' uitype='new_engagement' value='{"+engagementFromReq+"}' /> </filter></entity></fetch>";
                tracing.Trace("Start Plugin time:" +fetchQuery);

                EntityCollection engagements = service.RetrieveMultiple(new FetchExpression(fetchQuery));
                tracing.Trace("Count   :" + engagements.Entities.Count);

                
                tracing.Trace("Start Plugin time:" + DateTime.Now.ToString());
                MemoryStream ms = new MemoryStream();
                SpreadsheetDocument xl = SpreadsheetDocument.Create(ms, SpreadsheetDocumentType.Workbook);
                WorkbookPart wbp = xl.AddWorkbookPart();
                WorksheetPart wsp = null;
                try
                {
                    wsp = wbp.AddNewPart<WorksheetPart>();
                   
                }
                catch (Exception ex)
                {
                    throw new InvalidPluginExecutionException(ex.ToString());
                }
                Workbook wb = new Workbook();
                FileVersion fv = new FileVersion
                {
                    ApplicationName = "Microsoft Office Excel"
                };
                
                Worksheet ws = new Worksheet();

                //First cell
                SheetData sd = new SheetData();

                Columns columns = new Columns();
                columns.Append(CreateColumnData(1, 3, 23.5703125));
                columns.Append(CreateColumnData(4, 4, 11));
                columns.Append(CreateColumnData(6, 20, 18));
                ws.Append(columns);

                UInt32Value row_count = 1;
                Row row1 = new Row() { RowIndex = row_count };
                CreateHeader(service, row1, 1, ms, wb, xl, sd, ws, wbp, wsp, fv);
                int MainCountFlag = 1;
                int RowIndexCountFlag = 2;
                EntityCollection filtered = new EntityCollection();
                foreach (var engagement in engagements.Entities)
                {
                    try
                    {
                        if (((OptionSetValue)((AliasedValue)engagement.Attributes["new_confirmationrequest1_new_requesttypehandingfromclient"]).Value).Value.ToString() == reqTypeFromReq)
                        {
                            filtered.Entities.Add(engagement);
                        }
                    }
                    catch (Exception ex)
                    {

                    }

                }


                    foreach (var engagement in filtered.Entities)
                {


                    try
                    {
                        tracing.Trace(((AliasedValue)engagement.Attributes["new_confirmationrequest1_new_requesttypehandingfromclient"]).Value.ToString());
                        if (((OptionSetValue)((AliasedValue)engagement.Attributes["new_confirmationrequest1_new_requesttypehandingfromclient"]).Value).Value.ToString() == reqTypeFromReq)
                        {
                            //tracing.Trace("Afetr engagements : new_confirmationrequest1_new_confirmingparty :" + ((EntityReference)((AliasedValue)engagement.Attributes["new_confirmationrequest1_new_confirmingparty"]).Value).Name); ;
                            tracing.Trace("MainCountFlag :::::" + MainCountFlag);
                            engagementName = engagement.Attributes["new_name"].ToString();
                            var ss = new byte[] { };
                            row_count = Convert.ToUInt32(RowIndexCountFlag);// Convert.ToUInt32(MainCountFlag+1);
                            row1 = new Row() { RowIndex = Convert.ToUInt32(RowIndexCountFlag) };

                            Cell c1 = new Cell();
                            c1 = new Cell
                            {
                                CellReference = "A" + row_count,
                                DataType = CellValues.String,
                                StyleIndex = 2,
                                

                            CellValue = new CellValue(engagement.Attributes["new_name"].ToString())
                            };
                            row1.Append(c1);
                            tracing.Trace("first");

                            Cell c2 = new Cell();
                            c2 = new Cell
                            {
                                CellReference = "B" + row_count,
                                DataType = CellValues.String,
                                StyleIndex = 2,
                                CellValue = new CellValue((engagement.FormattedValues["new_confirmationrequest1_new_requesttypehandingfromclient"]))
                            };
                            row1.Append(c2);
                            tracing.Trace("first2");

                            Cell c3 = new Cell();
                            c3 = new Cell
                            {
                                CellReference = "C" + row_count,
                                DataType = CellValues.String,
                                StyleIndex = 2,
                                CellValue = new CellValue(((EntityReference)((AliasedValue)engagement.Attributes["new_confirmationrequest1_new_confirmingparty"]).Value).Name)
                            };
                            row1.Append(c3);
                            tracing.Trace("first3---" + (engagement.Attributes.ContainsKey("new_amount")));

                            Cell c4 = new Cell();
                            c4 = new Cell
                            {
                                CellReference = "D" + row_count,
                                DataType = CellValues.String,
                                StyleIndex = 2,
                                CellValue = engagement.Attributes.ContainsKey("new_confirmationrequest1_new_amount") ? new CellValue(engagement.FormattedValues["new_confirmationrequest1_new_amount"]) : null
                            };
                            row1.Append(c4);
                            tracing.Trace("first4");

                            Cell c5 = new Cell();
                            c5 = new Cell
                            {
                                CellReference = "E" + row_count,
                                DataType = CellValues.String,
                                StyleIndex = 2,
                                CellValue = engagement.Attributes.ContainsKey("new_confirmationrequest1_crf69_currencyname") ? new CellValue((engagement.FormattedValues["new_confirmationrequest1_crf69_currencyname"])) : null
                            };
                            row1.Append(c5);
                            tracing.Trace("first");

                            Cell c6 = new Cell();
                            c6 = new Cell
                            {
                                CellReference = "F" + row_count,
                                DataType = CellValues.String,
                                StyleIndex = 2,
                                CellValue = engagement.Attributes.ContainsKey("new_confirmationrequest1_crf69_portalamount") ? new CellValue(engagement.FormattedValues["new_confirmationrequest1_crf69_portalamount"].ToString()) : null
                            };
                            row1.Append(c6);
                            tracing.Trace("first6");

                            Cell c7 = new Cell();
                            c7 = new Cell
                            {
                                CellReference = "G" + row_count,
                                DataType = CellValues.String,
                                StyleIndex = 2,
                                CellValue = engagement.Attributes.ContainsKey("new_confirmationrequest1_crf69_portalcurrencies") ? new CellValue(engagement.FormattedValues["new_confirmationrequest1_crf69_portalcurrencies"].ToString()) : null
                            };
                            row1.Append(c7);
                            tracing.Trace("firssst7");

                            Cell c8 = new Cell();
                            c8 = new Cell
                            {
                                CellReference = "H" + row_count,
                                DataType = CellValues.String,
                                StyleIndex = 2,
                                CellValue = engagement.Attributes.ContainsKey("new_confirmationrequest1_new_comments") ? new CellValue(engagement.FormattedValues["new_confirmationrequest1_new_comments"].ToString()) : null
                            };
                            row1.Append(c8);
                            tracing.Trace("fsdirst8");

                            Cell c9 = new Cell();
                            c9 = new Cell
                            {
                                CellReference = "I" + row_count,
                                DataType = CellValues.String,
                                StyleIndex = 2,
                                CellValue = engagement.Attributes.ContainsKey("new_confirmationrequest1_crf69_clientcontactperson") ? new CellValue(engagement.FormattedValues["new_confirmationrequest1_crf69_clientcontactperson"].ToString()) : null
                            };
                            row1.Append(c9);
                            tracing.Trace("first9");

                            Cell c10 = new Cell();
                            c10 = new Cell
                            {
                                CellReference = "J" + row_count,
                                DataType = CellValues.String,
                                StyleIndex = 2,
                                CellValue = engagement.Attributes.ContainsKey("new_confirmationrequest1_crf69_confirmingpartycontact") ? new CellValue(engagement.FormattedValues["new_confirmationrequest1_crf69_confirmingpartycontact"].ToString()) : null
                            };
                            row1.Append(c10);
                            tracing.Trace("first10");

                            Cell c11 = new Cell();
                            c11 = new Cell
                            {
                                CellReference = "K" + row_count,
                                DataType = CellValues.String,
                                StyleIndex = 2,
                                CellValue = engagement.Attributes.ContainsKey("new_confirmationrequest1_crbcb_automaticreminderdate1") ? new CellValue(engagement.FormattedValues["new_confirmationrequest1_crbcb_automaticreminderdate1"].ToString()) : null
                            };
                            row1.Append(c11);
                            tracing.Trace("first11");

                            Cell c12 = new Cell();
                            c12 = new Cell
                            {
                                CellReference = "L" + row_count,
                                DataType = CellValues.String,
                                StyleIndex = 2,
                                CellValue = engagement.Attributes.ContainsKey("new_confirmationrequest1_crbcb_automaticreminderdate2") ? new CellValue(engagement.FormattedValues["new_confirmationrequest1_crbcb_automaticreminderdate2"].ToString()) : null
                            };
                            row1.Append(c12);
                            tracing.Trace("first12");

                            Cell c13 = new Cell();
                            c13 = new Cell
                            {
                                CellReference = "M" + row_count,
                                DataType = CellValues.String,
                                StyleIndex = 2,
                                CellValue = engagement.Attributes.ContainsKey("new_confirmationrequest1_crbcb_manualreminderdate") ? new CellValue(engagement.FormattedValues["new_confirmationrequest1_crbcb_manualreminderdate"].ToString()) : null
                            };
                            row1.Append(c13);

                            Cell c14 = new Cell();
                            c14 = new Cell
                            {
                                CellReference = "N" + row_count,
                                DataType = CellValues.String,
                                StyleIndex = 2,
                                CellValue = engagement.Attributes.ContainsKey("new_confirmationrequest1_crf69_clientcontactperson") ? new CellValue(engagement.FormattedValues["new_confirmationrequest1_crf69_clientcontactperson"].ToString()) : null
                            };
                            row1.Append(c14);

                            Cell c19 = new Cell();
                            c19 = new Cell
                            {
                                CellReference = "O" + row_count,
                                DataType = CellValues.String,
                                StyleIndex = 2,
                                CellValue = null
                            };
                            row1.Append(c19);
                            Cell c20 = new Cell();
                            c20 = new Cell
                            {
                                CellReference = "P" + row_count,
                                DataType = CellValues.String,
                                StyleIndex = 2,
                                CellValue = null
                            };
                            row1.Append(c20);

                            Cell c15 = new Cell();
                            c15 = new Cell
                            {
                                CellReference = "Q" + row_count,
                                DataType = CellValues.String,
                                StyleIndex = 1,
                                CellValue = engagement.Attributes.ContainsKey("new_confirmationrequest1_new_authletterrequestdate") ? new CellValue(engagement.FormattedValues["new_confirmationrequest1_new_authletterrequestdate"].ToString()) : null
                            };
                            row1.Append(c15);

                            Cell c16 = new Cell();
                            c16 = new Cell
                            {
                                CellReference = "R" + row_count,
                                DataType = CellValues.String,
                                StyleIndex = 1,
                                CellValue = engagement.Attributes.ContainsKey("new_confirmationrequest1_new_authletterapprovaldate") ? new CellValue(engagement.FormattedValues["new_confirmationrequest1_new_authletterapprovaldate"].ToString()) : null
                            };
                            row1.Append(c16);


                            Cell c17 = new Cell();
                            c17 = new Cell
                            {
                                CellReference = "S" + row_count,
                                DataType = CellValues.String,
                                StyleIndex = 1,
                                CellValue = engagement.Attributes.ContainsKey("new_confirmationrequest1_new_brachresponserequestdate") ? new CellValue(engagement.FormattedValues["new_confirmationrequest1_new_brachresponserequestdate"].ToString()) : null
                            };
                            row1.Append(c17);

                            Cell c18 = new Cell();
                            c18 = new Cell
                            {
                                CellReference = "T" + row_count,
                                DataType = CellValues.String,
                                StyleIndex = 1,
                                CellValue = engagement.Attributes.ContainsKey("new_confirmationrequest1_new_branchresponsereceiveddate") ? new CellValue(engagement.FormattedValues["new_confirmationrequest1_new_branchresponsereceiveddate"].ToString()) : null
                            };
                            row1.Append(c18);
                             ss = new byte[] { };
                            if (MainCountFlag == filtered.Entities.Count)
                            {
                                ss = CreateExcelDoc(service, row1, 1, ms, wb, xl, sd, ws, wbp, wsp, fv, true);


                            }
                            else
                            {
                                ss = CreateExcelDoc(service, row1, 1, ms, wb, xl, sd, ws, wbp, wsp, fv, false);

                            }


                            RowIndexCountFlag++;


                        }
                    }
                    catch (Exception ex)
                    {
                        tracing.Trace("Exception ss" + ex.Message);

                    }
                    MainCountFlag++;
                    
                    tracing.Trace("MainCountFlag ss" + MainCountFlag);
                    //if (MainCountFlag != 0)
                    

                }





            }
            catch (FaultException<OrganizationServiceFault> ex)
            {
                tracing.Trace("Error:" + ex);
                throw new InvalidPluginExecutionException("An error occurred.Pleaes try again.");
                tracing.Trace("Error:" + ex);
            }
        }

      
        public Entity getJournalAttachment(Guid summaryID, IOrganizationService service)
        {
            Entity annotationObj = service.Retrieve("new_summaryreportentry", summaryID, new ColumnSet(true));
            return annotationObj; ;
        }
      
        public byte[] CreateExcelDoc(IOrganizationService service, Row r1, UInt32Value rowcount, MemoryStream ms, Workbook wb, SpreadsheetDocument xl, SheetData sd, Worksheet ws, WorkbookPart wbp, WorksheetPart wsp, FileVersion fv, bool Upload)
        {
            sd.Append(r1);
            byte[] dt = null;
            if (Upload)
            {

                ws.Append(sd);
                wsp.Worksheet = ws;

                SheetProtection sheetProtection = new SheetProtection
                {
                    
                    // these are the "default" Excel settings when you do a normal protect
                    Sheet = true,
                    Objects = true,
                    Scenarios = true
                };
                sheetProtection.Password = new HexBinaryValue("1234");//"tR84v9D$!p";
                ProtectedRanges pRanges = new ProtectedRanges();
                ProtectedRange pRange = new ProtectedRange();
                ListValue<StringValue> lValue = new ListValue<StringValue>
                {
                    InnerText = "A:P" //set cell which you want to make it editable
                };
                pRange.SequenceOfReferences = lValue;
                pRange.Name = "not allow editing";
                pRanges.Append(pRange);              
                PageMargins pageM = wsp.Worksheet.GetFirstChild<PageMargins>();

                wsp.Worksheet.InsertBefore(sheetProtection, pageM);
                wsp.Worksheet.InsertBefore(pRanges, pageM);
                bool bFound = false;
                OpenXmlElement oxe = wsp.Worksheet.FirstChild;
                foreach (var child in wsp.Worksheet.ChildElements)
                {
                    // start with SheetData because it's a required child element
                    if (child is SheetData || child is SheetCalculationProperties)
                    {
                        tracing.Trace("oxe:" + child.XName.LocalName);

                        oxe = child;
                        bFound = true;
                    }
                }

                //if (bFound) wsp.Worksheet.InsertAfter(sheetProtection, oxe);
                //else wsp.Worksheet.PrependChild(sheetProtection);
                WorkbookStylesPart wbsp = wbp.AddNewPart<WorkbookStylesPart>();
                CellFormat lockFormat = new CellFormat()
                {
                    ApplyProtection = true,
                    Protection = new Protection()
                    {
                        Locked = true
                    }
                };


                // add styles to sheet
                wbsp.Stylesheet = CreateStylesheet(tracing);
                wbsp.Stylesheet.CellFormats.AppendChild<CellFormat>(lockFormat);
                wbsp.Stylesheet.CellFormats.Count = UInt32Value.FromUInt32((uint)wbsp.Stylesheet.CellFormats.ChildElements.Count);
                wbsp.Stylesheet.Save();
                
                wsp.Worksheet.Save();
                Sheets sheets = new Sheets();
                Sheet sheet = new Sheet
                {
                    Name = engagementName,
                    SheetId = 1,
                    Id = wbp.GetIdOfPart(wsp)
                };
     
                sheets.Append(sheet);
                wb.Append(fv);
                wb.Append(sheets);

                xl.WorkbookPart.Workbook = wb;
                xl.WorkbookPart.Workbook.Save();
                xl.Close();
                dt = ms.ToArray();


                Entity note = new
                Entity("annotation");

                note["subject"] = "Summary Report";

                note["filename"] = "SummaryReport.xlsx";

                note["documentbody"] = Convert.ToBase64String(dt);
                note["objectid"] = new EntityReference("new_summaryreportentry", SummaryId);
                var attachmentId = service.Create(note);

                Entity status = new
                Entity("new_summaryreportentry", SummaryId);
                status["new_reportcreated"] = "Yes";
                service.Update(status);
            }

            //}

            return dt;

        }

        private static Stylesheet CreateStylesheet(ITracingService tracing)
        {
            Stylesheet stylesheet1 = new Stylesheet() { MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "x14ac" } };
            stylesheet1.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            stylesheet1.AddNamespaceDeclaration("x14ac", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac");

            Fonts fonts1 = new Fonts() { Count = (UInt32Value)1U, KnownFonts = true };

            Font font1 = new Font();
            FontSize fontSize1 = new FontSize() { Val = 10D };

            Color color1 = new Color() { Theme = (UInt32Value)1U };
            FontName fontName1 = new FontName() { Val = "Arial" };
            FontFamilyNumbering fontFamilyNumbering1 = new FontFamilyNumbering() { Val = 2 };
            FontScheme fontScheme1 = new FontScheme() { Val = FontSchemeValues.Minor };

            Font font2 = new Font();
            FontSize fontSize2 = new FontSize() { Val = 10D };
            Color color2 = new Color() { Rgb = "FFFFFF" };
            FontName fontName2 = new FontName() { Val = "Arial" };
            FontFamilyNumbering fontFamilyNumbering2 = new FontFamilyNumbering() { Val = 2 };
            FontScheme fontScheme2 = new FontScheme() { Val = FontSchemeValues.Minor };

            font1.Append(fontSize1);
            font1.Append(color1);
            font1.Append(fontName1);
            font1.Append(fontFamilyNumbering1);
            font1.Append(fontScheme1);
            font2.Append(fontSize2);
            font2.Append(color2);
            font2.Append(fontName2);
            font2.Append(fontFamilyNumbering2);
            font2.Append(fontScheme2);
            font2.Append(new Bold());
            

            fonts1.Append(font1);
            fonts1.Append(font2);
            Fills fills1 = new Fills() { Count = (UInt32Value)5U };

            // FillId = 0
            //Fill fill1 = new Fill();
            //PatternFill patternFill1 = new PatternFill() { PatternType = PatternValues.None };
            //ForegroundColor foregroundColort = new ForegroundColor() { Rgb = "B4B3B1" };
            //BackgroundColor backgroundColort = new BackgroundColor() { Indexed = (UInt32Value)64U };
            //patternFill1.Append(foregroundColort);
            //patternFill1.Append(backgroundColort);
            //fill1.Append(patternFill1);
            // FillId = 1
            Fill fill1 = new Fill();
            PatternFill patternFill1 = new PatternFill() { PatternType = PatternValues.None };
            fill1.Append(patternFill1);

            // FillId = 1
            Fill fill2 = new Fill();
            PatternFill patternFill2 = new PatternFill() { PatternType = PatternValues.Gray125 };
            fill2.Append(patternFill2);

            // FillId = 2,light yellow
            Fill fill3 = new Fill();
            PatternFill patternFill3 = new PatternFill() { PatternType = PatternValues.Solid };
            ForegroundColor foregroundColor1 = new ForegroundColor() { Rgb = "B4B3B1" };
            BackgroundColor backgroundColor1 = new BackgroundColor() { Indexed = (UInt32Value)64U };
            patternFill3.Append(foregroundColor1);
            patternFill3.Append(backgroundColor1);
            fill3.Append(patternFill3);

            // FillId = 3,BLUE
            Fill fill4 = new Fill();
            PatternFill patternFill4 = new PatternFill() { PatternType = PatternValues.Solid };
            ForegroundColor foregroundColor2 = new ForegroundColor() { Rgb = "B4B3B1" };
            BackgroundColor backgroundColor2 = new BackgroundColor() { Indexed = (UInt32Value)64U };
            patternFill4.Append(foregroundColor2);
            patternFill4.Append(backgroundColor2);
            fill4.Append(patternFill4);

            // FillId = 4,blue actual
            Fill fill5 = new Fill();
            PatternFill patternFill5 = new PatternFill() { PatternType = PatternValues.Solid };
            ForegroundColor foregroundColor3 = new ForegroundColor() { Rgb = "7CA5DF" };
            BackgroundColor backgroundColor3 = new BackgroundColor() { Indexed = (UInt32Value)64U };
            patternFill5.Append(foregroundColor3);
            patternFill5.Append(backgroundColor3);
            fill5.Append(patternFill5);

            fills1.Append(fill1);
            fills1.Append(fill2);
            fills1.Append(fill3);
            fills1.Append(fill4);
            fills1.Append(fill5);

            Borders borders1 = new Borders() { Count = (UInt32Value)1U };

            Border border1 = new Border();
            DocumentFormat.OpenXml.Spreadsheet.LeftBorder leftBorder1 = new DocumentFormat.OpenXml.Spreadsheet.LeftBorder() { Style = DocumentFormat.OpenXml.Spreadsheet.BorderStyleValues.Thick };
            DocumentFormat.OpenXml.Spreadsheet.RightBorder rightBorder1 = new DocumentFormat.OpenXml.Spreadsheet.RightBorder() { Style = DocumentFormat.OpenXml.Spreadsheet.BorderStyleValues.Thick };
            DocumentFormat.OpenXml.Spreadsheet.TopBorder topBorder1 = new DocumentFormat.OpenXml.Spreadsheet.TopBorder() { Style = DocumentFormat.OpenXml.Spreadsheet.BorderStyleValues.Thick };
            DocumentFormat.OpenXml.Spreadsheet.BottomBorder bottomBorder1 = new DocumentFormat.OpenXml.Spreadsheet.BottomBorder() { Style = DocumentFormat.OpenXml.Spreadsheet.BorderStyleValues.Thick };
            DocumentFormat.OpenXml.Spreadsheet.DiagonalBorder diagonalBorder1 = new DocumentFormat.OpenXml.Spreadsheet.DiagonalBorder() { Style = DocumentFormat.OpenXml.Spreadsheet.BorderStyleValues.Thick };


            border1.Append(leftBorder1);
            border1.Append(rightBorder1);
            border1.Append(topBorder1);
            border1.Append(bottomBorder1);
            border1.Append(diagonalBorder1);

            borders1.Append(border1);
            tracing.Trace("oxe:" + borders1.Count);

            CellStyleFormats cellStyleFormats1 = new CellStyleFormats() { Count = (UInt32Value)1U };
            CellFormat cellFormat1 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U };

            cellStyleFormats1.Append(cellFormat1);

            CellFormats cellFormats1 = new CellFormats() { Count = (UInt32Value)4U };
            CellFormat cellFormat2 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U, ApplyProtection = true, Protection = new Protection() { Locked = true } };
            CellFormat cellFormat3 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)3U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U, ApplyFill = true, ApplyProtection = true, Protection = new Protection() { Locked = true } };
            CellFormat cellFormat4 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U, ApplyFill = true, ApplyProtection = true, Protection = new Protection() { Locked = true } };
            CellFormat cellFormat5 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)1U, FillId = (UInt32Value)4U, BorderId = 0, FormatId = (UInt32Value)0U, ApplyFill = true, ApplyProtection = true, Protection = new Protection() { Locked = true } };
            
            cellFormat5.Alignment = new Alignment() { WrapText = true, Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Center };
            cellFormats1.Append(cellFormat2);
            cellFormats1.Append(cellFormat3);
            cellFormats1.Append(cellFormat4);
            cellFormats1.Append(cellFormat5);

            CellStyles cellStyles1 = new CellStyles() { Count = (UInt32Value)1U };
            CellStyle cellStyle1 = new CellStyle() { Name = "Normal", FormatId = (UInt32Value)0U, BuiltinId = (UInt32Value)0U };

            cellStyles1.Append(cellStyle1);
            DifferentialFormats differentialFormats1 = new DifferentialFormats() { Count = (UInt32Value)0U };
            TableStyles tableStyles1 = new TableStyles() { Count = (UInt32Value)0U, DefaultTableStyle = "TableStyleMedium2", DefaultPivotStyle = "PivotStyleMedium9" };

            StylesheetExtensionList stylesheetExtensionList1 = new StylesheetExtensionList();

            StylesheetExtension stylesheetExtension1 = new StylesheetExtension() { Uri = "{EB79DEF2-80B8-43e5-95BD-54CBDDF9020C}" };
            stylesheetExtension1.AddNamespaceDeclaration("x14", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main");
            X14.SlicerStyles slicerStyles1 = new X14.SlicerStyles() { DefaultSlicerStyle = "SlicerStyleLight1" };

            stylesheetExtension1.Append(slicerStyles1);

            stylesheetExtensionList1.Append(stylesheetExtension1);

            stylesheet1.Append(fonts1);
            stylesheet1.Append(fills1);
            stylesheet1.Append(borders1);
            stylesheet1.Append(cellStyleFormats1);
            stylesheet1.Append(cellFormats1);
            stylesheet1.Append(cellStyles1);
            stylesheet1.Append(differentialFormats1);
            stylesheet1.Append(tableStyles1);
            stylesheet1.Append(stylesheetExtensionList1);
            return stylesheet1;
        }

        public void CreateHeader(IOrganizationService service, Row r1, UInt32Value rowcount, MemoryStream ms, Workbook wb, SpreadsheetDocument xl, SheetData sd, Worksheet ws, WorkbookPart wbp, WorksheetPart wsp, FileVersion fv)
        {
            Cell c1 = new Cell
            {
                CellReference = "A1",
                DataType = CellValues.String,
                CellValue = new CellValue("Engagement Name"),
                StyleIndex = (UInt32Value)3U
            };
            r1.Append(c1);

            // Second cell
            Cell c2 = new Cell
            {
                CellReference = "B1",
                DataType = CellValues.String,
                CellValue = new CellValue("Request type"),
                StyleIndex = (UInt32Value)3U
            };
            r1.Append(c2);

            // Third cell
            Cell c3 = new Cell
            {
                CellReference = "C1",
                DataType = CellValues.String,
                CellValue = new CellValue("Confirming Party"),
                StyleIndex = (UInt32Value)3U
            };
            r1.Append(c3);

            // Fourth cell
            Cell c4 = new Cell
            {
                CellReference = "D1",
                DataType = CellValues.String,
                CellValue = new CellValue("Amount"),
                StyleIndex = (UInt32Value)3U
            };
            r1.Append(c4);

            // Fifth cell
            Cell c5 = new Cell
            {
                CellReference = "E1",
                DataType = CellValues.String,
                CellValue = new CellValue("Currency Name"),
                StyleIndex = (UInt32Value)3U
            };
            r1.Append(c5);

            // Sixth cell
            Cell c6 = new Cell
            {
                CellReference = "F1",
                DataType = CellValues.String,
                CellValue = new CellValue("Portal Amount (Confirming Party)"),
                StyleIndex = (UInt32Value)3U
            };
            r1.Append(c6);

            // Seventh cell
            Cell c7 = new Cell
            {
                CellReference = "G1",
                DataType = CellValues.String,
                CellValue = new CellValue("Portal Currency (Confirming Party)"),
                StyleIndex = (UInt32Value)3U
            };
            r1.Append(c7);

            // Eighth cell
            Cell c8 = new Cell
            {
                CellReference = "H1",
                DataType = CellValues.String,
                CellValue = new CellValue("Portal Comments (Confirming Party )"),
                StyleIndex = (UInt32Value)3U
            };
            r1.Append(c8);

            // Ninth cell
            Cell c9 = new Cell
            {
                CellReference = "I1",
                DataType = CellValues.String,
                CellValue = new CellValue("Full Name (Client Contact)"),
                StyleIndex = (UInt32Value)3U
            };
            r1.Append(c9);

            

            // Eleventh cell
            Cell c11 = new Cell
            {
                CellReference = "J1",
                DataType = CellValues.String,
                StyleIndex = (UInt32Value)3U,
                CellValue = new CellValue("Confirming Party Contact")
            };
            r1.Append(c11);

            // Twelveth cell
            Cell c12 = new Cell
            {
                CellReference = "K1",
                StyleIndex = (UInt32Value)3U,
                DataType = CellValues.String,
                CellValue = new CellValue("AutomaticReminderDate1")
            };
            r1.Append(c12);

            // Thirteenth cell
            Cell c13 = new Cell
            {
                CellReference = "L1",
                DataType = CellValues.String,
                StyleIndex = (UInt32Value)3U,
                CellValue = new CellValue("AutomaticReminderDate2")
            };
            r1.Append(c13);


            // Fourteenth cell
            Cell c14 = new Cell
            {
                CellReference = "M1",
                DataType = CellValues.String,
                StyleIndex = (UInt32Value)3U,

                CellValue = new CellValue("ManualReminderDate")
            };
            r1.Append(c14);


            // Fifteenth cell
            Cell c15 = new Cell
            {
                CellReference = "N1"
            };
            c15.DataType = CellValues.String;
            c15.StyleIndex = (UInt32Value)3U;
            c15.CellValue = new CellValue("Client Contact Person");
            r1.Append(c15);

            Cell c21 = new Cell
            {
                CellReference = "O1",
                DataType = CellValues.String,
                StyleIndex = (UInt32Value)3U,
                CellValue = new CellValue("Reconciliations, as appropriate")
            };
            r1.Append(c21);
            Cell c19 = new Cell
            {
                CellReference = "P1",
                DataType = CellValues.String,
                StyleIndex = (UInt32Value)3U,
                CellValue = new CellValue("Alternative / additional procedures, as appropriate")
            };
            r1.Append(c19);

            // Sizteenth cell
            Cell c16 = new Cell
            {
                CellReference = "Q1",
                DataType = CellValues.String,
                StyleIndex = (UInt32Value)3U,
                CellValue = new CellValue("Date sent to client for authorization ")
            };
            r1.Append(c16);


            // Sizteenth cell
            Cell c17 = new Cell
            {
                CellReference = "R1",
                DataType = CellValues.String,
                StyleIndex = (UInt32Value)3U,
                CellValue = new CellValue("Authorization Letter Approval Date")
            };
            r1.Append(c17);


            // Seventeen cell
            Cell c18 = new Cell
            {
                CellReference = "S1",
                DataType = CellValues.String,
                StyleIndex = (UInt32Value)3U,
                CellValue = new CellValue("Confirmation Response Request Date")
            };
            r1.Append(c18);
            // Seventeen cell
            Cell c20 = new Cell
            {
                CellReference = "T1",
                DataType = CellValues.String,
                StyleIndex = (UInt32Value)3U,
                CellValue = new CellValue("Confirmation Response Received Date")
            };
            r1.Append(c20);
           




            sd.Append(r1);

            
        }

        private static Column CreateColumnData(UInt32 StartColumnIndex, UInt32 EndColumnIndex, double ColumnWidth)
        {
            Column column;
            column = new Column();
            column.Min = StartColumnIndex;
            column.Max = EndColumnIndex;
            column.Width = ColumnWidth;
            column.CustomWidth = true;
            return column;
        }

    }

}
