using DocumentFormat.OpenXml.Packaging;
using Ap = DocumentFormat.OpenXml.ExtendedProperties;
using Vt = DocumentFormat.OpenXml.VariantTypes;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using X14 = DocumentFormat.OpenXml.Office2010.Excel;
using A = DocumentFormat.OpenXml.Drawing;


namespace Cassini.UI.Service
{
    public class XMLCreator
    {
        
            // Creates a SpreadsheetDocument.
            public void CreatePackage(string filePath)
            {
                using (SpreadsheetDocument package = SpreadsheetDocument.Create(filePath, SpreadsheetDocumentType.Workbook))
                {
                    CreateParts(package);
                }
            }

            // Adds child parts and generates content of the specified part.
            private void CreateParts(SpreadsheetDocument document)
            {
                ExtendedFilePropertiesPart extendedFilePropertiesPart1 = document.AddNewPart<ExtendedFilePropertiesPart>("rId3");
                GenerateExtendedFilePropertiesPart1Content(extendedFilePropertiesPart1);

                WorkbookPart workbookPart1 = document.AddWorkbookPart();
                GenerateWorkbookPart1Content(workbookPart1);

                PivotTableCacheDefinitionPart pivotTableCacheDefinitionPart1 = workbookPart1.AddNewPart<PivotTableCacheDefinitionPart>("rId3");
                GeneratePivotTableCacheDefinitionPart1Content(pivotTableCacheDefinitionPart1);

                PivotTableCacheRecordsPart pivotTableCacheRecordsPart1 = pivotTableCacheDefinitionPart1.AddNewPart<PivotTableCacheRecordsPart>("rId1");
                GeneratePivotTableCacheRecordsPart1Content(pivotTableCacheRecordsPart1);

                CalculationChainPart calculationChainPart1 = workbookPart1.AddNewPart<CalculationChainPart>("rId7");
                GenerateCalculationChainPart1Content(calculationChainPart1);

                WorksheetPart worksheetPart1 = workbookPart1.AddNewPart<WorksheetPart>("rId2");
                GenerateWorksheetPart1Content(worksheetPart1);

                TableDefinitionPart tableDefinitionPart1 = worksheetPart1.AddNewPart<TableDefinitionPart>("rId1");
                GenerateTableDefinitionPart1Content(tableDefinitionPart1);

                WorksheetPart worksheetPart2 = workbookPart1.AddNewPart<WorksheetPart>("rId1");
                GenerateWorksheetPart2Content(worksheetPart2);

                PivotTablePart pivotTablePart1 = worksheetPart2.AddNewPart<PivotTablePart>("rId1");
                GeneratePivotTablePart1Content(pivotTablePart1);

                pivotTablePart1.AddPart(pivotTableCacheDefinitionPart1, "rId1");

                SharedStringTablePart sharedStringTablePart1 = workbookPart1.AddNewPart<SharedStringTablePart>("rId6");
                GenerateSharedStringTablePart1Content(sharedStringTablePart1);

                WorkbookStylesPart workbookStylesPart1 = workbookPart1.AddNewPart<WorkbookStylesPart>("rId5");
                GenerateWorkbookStylesPart1Content(workbookStylesPart1);

                ThemePart themePart1 = workbookPart1.AddNewPart<ThemePart>("rId4");
                GenerateThemePart1Content(themePart1);

                SetPackageProperties(document);
            }

            // Generates content of extendedFilePropertiesPart1.
            private void GenerateExtendedFilePropertiesPart1Content(ExtendedFilePropertiesPart extendedFilePropertiesPart1)
            {
                Ap.Properties properties1 = new Ap.Properties();
                properties1.AddNamespaceDeclaration("vt", "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes");
                Ap.Application application1 = new Ap.Application();
                application1.Text = "Microsoft Excel";
                Ap.DocumentSecurity documentSecurity1 = new Ap.DocumentSecurity();
                documentSecurity1.Text = "0";
                Ap.ScaleCrop scaleCrop1 = new Ap.ScaleCrop();
                scaleCrop1.Text = "false";

                Ap.HeadingPairs headingPairs1 = new Ap.HeadingPairs();

                Vt.VTVector vTVector1 = new Vt.VTVector() { BaseType = Vt.VectorBaseValues.Variant, Size = (UInt32Value)2U };

                Vt.Variant variant1 = new Vt.Variant();
                Vt.VTLPSTR vTLPSTR1 = new Vt.VTLPSTR();
                vTLPSTR1.Text = "Листы";

                variant1.Append(vTLPSTR1);

                Vt.Variant variant2 = new Vt.Variant();
                Vt.VTInt32 vTInt321 = new Vt.VTInt32();
                vTInt321.Text = "2";

                variant2.Append(vTInt321);

                vTVector1.Append(variant1);
                vTVector1.Append(variant2);

                headingPairs1.Append(vTVector1);

                Ap.TitlesOfParts titlesOfParts1 = new Ap.TitlesOfParts();

                Vt.VTVector vTVector2 = new Vt.VTVector() { BaseType = Vt.VectorBaseValues.Lpstr, Size = (UInt32Value)2U };
                Vt.VTLPSTR vTLPSTR2 = new Vt.VTLPSTR();
                vTLPSTR2.Text = "Лист1";
                Vt.VTLPSTR vTLPSTR3 = new Vt.VTLPSTR();
                vTLPSTR3.Text = "Марусинець_Мейсарош_Гобан_СПД_0";

                vTVector2.Append(vTLPSTR2);
                vTVector2.Append(vTLPSTR3);

                titlesOfParts1.Append(vTVector2);
                Ap.LinksUpToDate linksUpToDate1 = new Ap.LinksUpToDate();
                linksUpToDate1.Text = "false";
                Ap.SharedDocument sharedDocument1 = new Ap.SharedDocument();
                sharedDocument1.Text = "false";
                Ap.HyperlinksChanged hyperlinksChanged1 = new Ap.HyperlinksChanged();
                hyperlinksChanged1.Text = "false";
                Ap.ApplicationVersion applicationVersion1 = new Ap.ApplicationVersion();
                applicationVersion1.Text = "16.0300";

                properties1.Append(application1);
                properties1.Append(documentSecurity1);
                properties1.Append(scaleCrop1);
                properties1.Append(headingPairs1);
                properties1.Append(titlesOfParts1);
                properties1.Append(linksUpToDate1);
                properties1.Append(sharedDocument1);
                properties1.Append(hyperlinksChanged1);
                properties1.Append(applicationVersion1);

                extendedFilePropertiesPart1.Properties = properties1;
            }

            // Generates content of workbookPart1.
            private void GenerateWorkbookPart1Content(WorkbookPart workbookPart1)
            {
                Workbook workbook1 = new Workbook() { MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "x15" } };
                workbook1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
                workbook1.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
                workbook1.AddNamespaceDeclaration("x15", "http://schemas.microsoft.com/office/spreadsheetml/2010/11/main");
                FileVersion fileVersion1 = new FileVersion() { ApplicationName = "xl", LastEdited = "6", LowestEdited = "6", BuildVersion = "14420" };
                WorkbookProperties workbookProperties1 = new WorkbookProperties() { DefaultThemeVersion = (UInt32Value)164011U };

                AlternateContent alternateContent1 = new AlternateContent();
                alternateContent1.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");

                AlternateContentChoice alternateContentChoice1 = new AlternateContentChoice() { Requires = "x15" };

                OpenXmlUnknownElement openXmlUnknownElement1 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<x15ac:absPath xmlns:x15ac=\"http://schemas.microsoft.com/office/spreadsheetml/2010/11/ac\" url=\"C:\\Users\\bezvershuk_do.ORANTA\\Desktop\\\" />");

                alternateContentChoice1.Append(openXmlUnknownElement1);

                alternateContent1.Append(alternateContentChoice1);

                BookViews bookViews1 = new BookViews();
                WorkbookView workbookView1 = new WorkbookView() { XWindow = 0, YWindow = 0, WindowWidth = (UInt32Value)28800U, WindowHeight = (UInt32Value)12885U, ActiveTab = (UInt32Value)1U };

                bookViews1.Append(workbookView1);

                Sheets sheets1 = new Sheets();
                Sheet sheet1 = new Sheet() { Name = "Лист1", SheetId = (UInt32Value)2U, Id = "rId1" };
                Sheet sheet2 = new Sheet() { Name = "Марусинець_Мейсарош_Гобан_СПД_0", SheetId = (UInt32Value)1U, Id = "rId2" };

                sheets1.Append(sheet1);
                sheets1.Append(sheet2);
                CalculationProperties calculationProperties1 = new CalculationProperties() { CalculationId = (UInt32Value)162913U };

                PivotCaches pivotCaches1 = new PivotCaches();
                PivotCache pivotCache1 = new PivotCache() { CacheId = (UInt32Value)1U, Id = "rId3" };

                pivotCaches1.Append(pivotCache1);

                workbook1.Append(fileVersion1);
                workbook1.Append(workbookProperties1);
                workbook1.Append(alternateContent1);
                workbook1.Append(bookViews1);
                workbook1.Append(sheets1);
                workbook1.Append(calculationProperties1);
                workbook1.Append(pivotCaches1);

                workbookPart1.Workbook = workbook1;
            }

            // Generates content of pivotTableCacheDefinitionPart1.
            private void GeneratePivotTableCacheDefinitionPart1Content(PivotTableCacheDefinitionPart pivotTableCacheDefinitionPart1)
            {
                PivotCacheDefinition pivotCacheDefinition1 = new PivotCacheDefinition() { Id = "rId1", RefreshedBy = "Безвершук Дмитро Олександрович", RefreshedDate = 42849.415533333333D, CreatedVersion = 6, RefreshedVersion = 6, MinRefreshableVersion = 3, RecordCount = (UInt32Value)13U };
                pivotCacheDefinition1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");

                CacheSource cacheSource1 = new CacheSource() { Type = SourceValues.Worksheet };
                WorksheetSource worksheetSource1 = new WorksheetSource() { Name = "Таблица1" };

                cacheSource1.Append(worksheetSource1);

                CacheFields cacheFields1 = new CacheFields() { Count = (UInt32Value)10U };

                CacheField cacheField1 = new CacheField() { Name = "ІПН", NumberFormatId = (UInt32Value)49U };

                SharedItems sharedItems1 = new SharedItems() { Count = (UInt32Value)3U };
                StringItem stringItem1 = new StringItem() { Val = "2844613435" };
                StringItem stringItem2 = new StringItem() { Val = "3425509799" };
                StringItem stringItem3 = new StringItem() { Val = "2209911436" };

                sharedItems1.Append(stringItem1);
                sharedItems1.Append(stringItem2);
                sharedItems1.Append(stringItem3);

                cacheField1.Append(sharedItems1);

                CacheField cacheField2 = new CacheField() { Name = "Агент", NumberFormatId = (UInt32Value)0U };

                SharedItems sharedItems2 = new SharedItems() { Count = (UInt32Value)3U };
                StringItem stringItem4 = new StringItem() { Val = "Марусинець З.С." };
                StringItem stringItem5 = new StringItem() { Val = "Мейсарош Е.Т." };
                StringItem stringItem6 = new StringItem() { Val = "Гобан Ю.Ю." };

                sharedItems2.Append(stringItem4);
                sharedItems2.Append(stringItem5);
                sharedItems2.Append(stringItem6);

                cacheField2.Append(sharedItems2);

                CacheField cacheField3 = new CacheField() { Name = "Код програми", NumberFormatId = (UInt32Value)49U };
                SharedItems sharedItems3 = new SharedItems();

                cacheField3.Append(sharedItems3);

                CacheField cacheField4 = new CacheField() { Name = "СП", NumberFormatId = (UInt32Value)0U };
                SharedItems sharedItems4 = new SharedItems() { ContainsSemiMixedTypes = false, ContainsString = false, ContainsNumber = true, ContainsInteger = true, MinValue = 697D, MaxValue = 7987D };

                cacheField4.Append(sharedItems4);

                CacheField cacheField5 = new CacheField() { Name = "АВ", NumberFormatId = (UInt32Value)0U };
                SharedItems sharedItems5 = new SharedItems() { ContainsSemiMixedTypes = false, ContainsString = false, ContainsNumber = true, MinValue = 139.4D, MaxValue = 1597.4D };

                cacheField5.Append(sharedItems5);

                CacheField cacheField6 = new CacheField() { Name = "Код відділення", NumberFormatId = (UInt32Value)49U };
                SharedItems sharedItems6 = new SharedItems();

                cacheField6.Append(sharedItems6);

                CacheField cacheField7 = new CacheField() { Name = "Канал", NumberFormatId = (UInt32Value)0U };

                SharedItems sharedItems7 = new SharedItems() { Count = (UInt32Value)1U };
                StringItem stringItem7 = new StringItem() { Val = "22 - фізичні особи - суб’єкти підприємницької діяльності (Категорія 1)" };

                sharedItems7.Append(stringItem7);

                cacheField7.Append(sharedItems7);

                CacheField cacheField8 = new CacheField() { Name = "Договір", NumberFormatId = (UInt32Value)0U };
                SharedItems sharedItems8 = new SharedItems();

                cacheField8.Append(sharedItems8);

                CacheField cacheField9 = new CacheField() { Name = "ID акту", NumberFormatId = (UInt32Value)49U };
                SharedItems sharedItems9 = new SharedItems();

                cacheField9.Append(sharedItems9);

                CacheField cacheField10 = new CacheField() { Name = "Дирекція", NumberFormatId = (UInt32Value)0U };

                SharedItems sharedItems10 = new SharedItems() { Count = (UInt32Value)1U };
                StringItem stringItem8 = new StringItem() { Val = "07" };

                sharedItems10.Append(stringItem8);

                cacheField10.Append(sharedItems10);

                cacheFields1.Append(cacheField1);
                cacheFields1.Append(cacheField2);
                cacheFields1.Append(cacheField3);
                cacheFields1.Append(cacheField4);
                cacheFields1.Append(cacheField5);
                cacheFields1.Append(cacheField6);
                cacheFields1.Append(cacheField7);
                cacheFields1.Append(cacheField8);
                cacheFields1.Append(cacheField9);
                cacheFields1.Append(cacheField10);

                PivotCacheDefinitionExtensionList pivotCacheDefinitionExtensionList1 = new PivotCacheDefinitionExtensionList();

                PivotCacheDefinitionExtension pivotCacheDefinitionExtension1 = new PivotCacheDefinitionExtension() { Uri = "{725AE2AE-9491-48be-B2B4-4EB974FC3084}" };
                pivotCacheDefinitionExtension1.AddNamespaceDeclaration("x14", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main");
                X14.PivotCacheDefinition pivotCacheDefinition2 = new X14.PivotCacheDefinition();

                pivotCacheDefinitionExtension1.Append(pivotCacheDefinition2);

                pivotCacheDefinitionExtensionList1.Append(pivotCacheDefinitionExtension1);

                pivotCacheDefinition1.Append(cacheSource1);
                pivotCacheDefinition1.Append(cacheFields1);
                pivotCacheDefinition1.Append(pivotCacheDefinitionExtensionList1);

                pivotTableCacheDefinitionPart1.PivotCacheDefinition = pivotCacheDefinition1;
            }

            // Generates content of pivotTableCacheRecordsPart1.
            private void GeneratePivotTableCacheRecordsPart1Content(PivotTableCacheRecordsPart pivotTableCacheRecordsPart1)
            {
                PivotCacheRecords pivotCacheRecords1 = new PivotCacheRecords() { Count = (UInt32Value)13U };
                pivotCacheRecords1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");

                PivotCacheRecord pivotCacheRecord1 = new PivotCacheRecord();
                FieldItem fieldItem1 = new FieldItem() { Val = (UInt32Value)0U };
                FieldItem fieldItem2 = new FieldItem() { Val = (UInt32Value)0U };
                StringItem stringItem9 = new StringItem() { Val = "231" };
                NumberItem numberItem1 = new NumberItem() { Val = 2917D };
                NumberItem numberItem2 = new NumberItem() { Val = 583.4D };
                StringItem stringItem10 = new StringItem() { Val = "0701" };
                FieldItem fieldItem3 = new FieldItem() { Val = (UInt32Value)0U };
                StringItem stringItem11 = new StringItem() { Val = "Агентський договір" };
                StringItem stringItem12 = new StringItem() { Val = "501898" };
                FieldItem fieldItem4 = new FieldItem() { Val = (UInt32Value)0U };

                pivotCacheRecord1.Append(fieldItem1);
                pivotCacheRecord1.Append(fieldItem2);
                pivotCacheRecord1.Append(stringItem9);
                pivotCacheRecord1.Append(numberItem1);
                pivotCacheRecord1.Append(numberItem2);
                pivotCacheRecord1.Append(stringItem10);
                pivotCacheRecord1.Append(fieldItem3);
                pivotCacheRecord1.Append(stringItem11);
                pivotCacheRecord1.Append(stringItem12);
                pivotCacheRecord1.Append(fieldItem4);

                PivotCacheRecord pivotCacheRecord2 = new PivotCacheRecord();
                FieldItem fieldItem5 = new FieldItem() { Val = (UInt32Value)0U };
                FieldItem fieldItem6 = new FieldItem() { Val = (UInt32Value)0U };
                StringItem stringItem13 = new StringItem() { Val = "232" };
                NumberItem numberItem3 = new NumberItem() { Val = 2517D };
                NumberItem numberItem4 = new NumberItem() { Val = 503.4D };
                StringItem stringItem14 = new StringItem() { Val = "0701" };
                FieldItem fieldItem7 = new FieldItem() { Val = (UInt32Value)0U };
                StringItem stringItem15 = new StringItem() { Val = "Агентський договір" };
                StringItem stringItem16 = new StringItem() { Val = "501898" };
                FieldItem fieldItem8 = new FieldItem() { Val = (UInt32Value)0U };

                pivotCacheRecord2.Append(fieldItem5);
                pivotCacheRecord2.Append(fieldItem6);
                pivotCacheRecord2.Append(stringItem13);
                pivotCacheRecord2.Append(numberItem3);
                pivotCacheRecord2.Append(numberItem4);
                pivotCacheRecord2.Append(stringItem14);
                pivotCacheRecord2.Append(fieldItem7);
                pivotCacheRecord2.Append(stringItem15);
                pivotCacheRecord2.Append(stringItem16);
                pivotCacheRecord2.Append(fieldItem8);

                PivotCacheRecord pivotCacheRecord3 = new PivotCacheRecord();
                FieldItem fieldItem9 = new FieldItem() { Val = (UInt32Value)0U };
                FieldItem fieldItem10 = new FieldItem() { Val = (UInt32Value)0U };
                StringItem stringItem17 = new StringItem() { Val = "231" };
                NumberItem numberItem5 = new NumberItem() { Val = 2091D };
                NumberItem numberItem6 = new NumberItem() { Val = 418.2D };
                StringItem stringItem18 = new StringItem() { Val = "0703" };
                FieldItem fieldItem11 = new FieldItem() { Val = (UInt32Value)0U };
                StringItem stringItem19 = new StringItem() { Val = "Агентський договір" };
                StringItem stringItem20 = new StringItem() { Val = "501898" };
                FieldItem fieldItem12 = new FieldItem() { Val = (UInt32Value)0U };

                pivotCacheRecord3.Append(fieldItem9);
                pivotCacheRecord3.Append(fieldItem10);
                pivotCacheRecord3.Append(stringItem17);
                pivotCacheRecord3.Append(numberItem5);
                pivotCacheRecord3.Append(numberItem6);
                pivotCacheRecord3.Append(stringItem18);
                pivotCacheRecord3.Append(fieldItem11);
                pivotCacheRecord3.Append(stringItem19);
                pivotCacheRecord3.Append(stringItem20);
                pivotCacheRecord3.Append(fieldItem12);

                PivotCacheRecord pivotCacheRecord4 = new PivotCacheRecord();
                FieldItem fieldItem13 = new FieldItem() { Val = (UInt32Value)0U };
                FieldItem fieldItem14 = new FieldItem() { Val = (UInt32Value)0U };
                StringItem stringItem21 = new StringItem() { Val = "231" };
                NumberItem numberItem7 = new NumberItem() { Val = 697D };
                NumberItem numberItem8 = new NumberItem() { Val = 139.4D };
                StringItem stringItem22 = new StringItem() { Val = "0704" };
                FieldItem fieldItem15 = new FieldItem() { Val = (UInt32Value)0U };
                StringItem stringItem23 = new StringItem() { Val = "Агентський договір" };
                StringItem stringItem24 = new StringItem() { Val = "501898" };
                FieldItem fieldItem16 = new FieldItem() { Val = (UInt32Value)0U };

                pivotCacheRecord4.Append(fieldItem13);
                pivotCacheRecord4.Append(fieldItem14);
                pivotCacheRecord4.Append(stringItem21);
                pivotCacheRecord4.Append(numberItem7);
                pivotCacheRecord4.Append(numberItem8);
                pivotCacheRecord4.Append(stringItem22);
                pivotCacheRecord4.Append(fieldItem15);
                pivotCacheRecord4.Append(stringItem23);
                pivotCacheRecord4.Append(stringItem24);
                pivotCacheRecord4.Append(fieldItem16);

                PivotCacheRecord pivotCacheRecord5 = new PivotCacheRecord();
                FieldItem fieldItem17 = new FieldItem() { Val = (UInt32Value)0U };
                FieldItem fieldItem18 = new FieldItem() { Val = (UInt32Value)0U };
                StringItem stringItem25 = new StringItem() { Val = "231" };
                NumberItem numberItem9 = new NumberItem() { Val = 6102D };
                NumberItem numberItem10 = new NumberItem() { Val = 1220.4000000000001D };
                StringItem stringItem26 = new StringItem() { Val = "0706" };
                FieldItem fieldItem19 = new FieldItem() { Val = (UInt32Value)0U };
                StringItem stringItem27 = new StringItem() { Val = "Агентський договір" };
                StringItem stringItem28 = new StringItem() { Val = "501898" };
                FieldItem fieldItem20 = new FieldItem() { Val = (UInt32Value)0U };

                pivotCacheRecord5.Append(fieldItem17);
                pivotCacheRecord5.Append(fieldItem18);
                pivotCacheRecord5.Append(stringItem25);
                pivotCacheRecord5.Append(numberItem9);
                pivotCacheRecord5.Append(numberItem10);
                pivotCacheRecord5.Append(stringItem26);
                pivotCacheRecord5.Append(fieldItem19);
                pivotCacheRecord5.Append(stringItem27);
                pivotCacheRecord5.Append(stringItem28);
                pivotCacheRecord5.Append(fieldItem20);

                PivotCacheRecord pivotCacheRecord6 = new PivotCacheRecord();
                FieldItem fieldItem21 = new FieldItem() { Val = (UInt32Value)0U };
                FieldItem fieldItem22 = new FieldItem() { Val = (UInt32Value)0U };
                StringItem stringItem29 = new StringItem() { Val = "231" };
                NumberItem numberItem11 = new NumberItem() { Val = 697D };
                NumberItem numberItem12 = new NumberItem() { Val = 139.4D };
                StringItem stringItem30 = new StringItem() { Val = "0708" };
                FieldItem fieldItem23 = new FieldItem() { Val = (UInt32Value)0U };
                StringItem stringItem31 = new StringItem() { Val = "Агентський договір" };
                StringItem stringItem32 = new StringItem() { Val = "501898" };
                FieldItem fieldItem24 = new FieldItem() { Val = (UInt32Value)0U };

                pivotCacheRecord6.Append(fieldItem21);
                pivotCacheRecord6.Append(fieldItem22);
                pivotCacheRecord6.Append(stringItem29);
                pivotCacheRecord6.Append(numberItem11);
                pivotCacheRecord6.Append(numberItem12);
                pivotCacheRecord6.Append(stringItem30);
                pivotCacheRecord6.Append(fieldItem23);
                pivotCacheRecord6.Append(stringItem31);
                pivotCacheRecord6.Append(stringItem32);
                pivotCacheRecord6.Append(fieldItem24);

                PivotCacheRecord pivotCacheRecord7 = new PivotCacheRecord();
                FieldItem fieldItem25 = new FieldItem() { Val = (UInt32Value)0U };
                FieldItem fieldItem26 = new FieldItem() { Val = (UInt32Value)0U };
                StringItem stringItem33 = new StringItem() { Val = "231" };
                NumberItem numberItem13 = new NumberItem() { Val = 6102D };
                NumberItem numberItem14 = new NumberItem() { Val = 1220.4000000000001D };
                StringItem stringItem34 = new StringItem() { Val = "0711" };
                FieldItem fieldItem27 = new FieldItem() { Val = (UInt32Value)0U };
                StringItem stringItem35 = new StringItem() { Val = "Агентський договір" };
                StringItem stringItem36 = new StringItem() { Val = "501898" };
                FieldItem fieldItem28 = new FieldItem() { Val = (UInt32Value)0U };

                pivotCacheRecord7.Append(fieldItem25);
                pivotCacheRecord7.Append(fieldItem26);
                pivotCacheRecord7.Append(stringItem33);
                pivotCacheRecord7.Append(numberItem13);
                pivotCacheRecord7.Append(numberItem14);
                pivotCacheRecord7.Append(stringItem34);
                pivotCacheRecord7.Append(fieldItem27);
                pivotCacheRecord7.Append(stringItem35);
                pivotCacheRecord7.Append(stringItem36);
                pivotCacheRecord7.Append(fieldItem28);

                PivotCacheRecord pivotCacheRecord8 = new PivotCacheRecord();
                FieldItem fieldItem29 = new FieldItem() { Val = (UInt32Value)1U };
                FieldItem fieldItem30 = new FieldItem() { Val = (UInt32Value)1U };
                StringItem stringItem37 = new StringItem() { Val = "231" };
                NumberItem numberItem15 = new NumberItem() { Val = 1807D };
                NumberItem numberItem16 = new NumberItem() { Val = 361.4D };
                StringItem stringItem38 = new StringItem() { Val = "0712" };
                FieldItem fieldItem31 = new FieldItem() { Val = (UInt32Value)0U };
                StringItem stringItem39 = new StringItem() { Val = "Агентський договір" };
                StringItem stringItem40 = new StringItem() { Val = "501897" };
                FieldItem fieldItem32 = new FieldItem() { Val = (UInt32Value)0U };

                pivotCacheRecord8.Append(fieldItem29);
                pivotCacheRecord8.Append(fieldItem30);
                pivotCacheRecord8.Append(stringItem37);
                pivotCacheRecord8.Append(numberItem15);
                pivotCacheRecord8.Append(numberItem16);
                pivotCacheRecord8.Append(stringItem38);
                pivotCacheRecord8.Append(fieldItem31);
                pivotCacheRecord8.Append(stringItem39);
                pivotCacheRecord8.Append(stringItem40);
                pivotCacheRecord8.Append(fieldItem32);

                PivotCacheRecord pivotCacheRecord9 = new PivotCacheRecord();
                FieldItem fieldItem33 = new FieldItem() { Val = (UInt32Value)2U };
                FieldItem fieldItem34 = new FieldItem() { Val = (UInt32Value)2U };
                StringItem stringItem41 = new StringItem() { Val = "231" };
                NumberItem numberItem17 = new NumberItem() { Val = 7987D };
                NumberItem numberItem18 = new NumberItem() { Val = 1597.4D };
                StringItem stringItem42 = new StringItem() { Val = "0715" };
                FieldItem fieldItem35 = new FieldItem() { Val = (UInt32Value)0U };
                StringItem stringItem43 = new StringItem() { Val = "Агентський договір" };
                StringItem stringItem44 = new StringItem() { Val = "501896" };
                FieldItem fieldItem36 = new FieldItem() { Val = (UInt32Value)0U };

                pivotCacheRecord9.Append(fieldItem33);
                pivotCacheRecord9.Append(fieldItem34);
                pivotCacheRecord9.Append(stringItem41);
                pivotCacheRecord9.Append(numberItem17);
                pivotCacheRecord9.Append(numberItem18);
                pivotCacheRecord9.Append(stringItem42);
                pivotCacheRecord9.Append(fieldItem35);
                pivotCacheRecord9.Append(stringItem43);
                pivotCacheRecord9.Append(stringItem44);
                pivotCacheRecord9.Append(fieldItem36);

                PivotCacheRecord pivotCacheRecord10 = new PivotCacheRecord();
                FieldItem fieldItem37 = new FieldItem() { Val = (UInt32Value)0U };
                FieldItem fieldItem38 = new FieldItem() { Val = (UInt32Value)0U };
                StringItem stringItem45 = new StringItem() { Val = "231" };
                NumberItem numberItem19 = new NumberItem() { Val = 697D };
                NumberItem numberItem20 = new NumberItem() { Val = 139.4D };
                StringItem stringItem46 = new StringItem() { Val = "0716" };
                FieldItem fieldItem39 = new FieldItem() { Val = (UInt32Value)0U };
                StringItem stringItem47 = new StringItem() { Val = "Агентський договір" };
                StringItem stringItem48 = new StringItem() { Val = "501898" };
                FieldItem fieldItem40 = new FieldItem() { Val = (UInt32Value)0U };

                pivotCacheRecord10.Append(fieldItem37);
                pivotCacheRecord10.Append(fieldItem38);
                pivotCacheRecord10.Append(stringItem45);
                pivotCacheRecord10.Append(numberItem19);
                pivotCacheRecord10.Append(numberItem20);
                pivotCacheRecord10.Append(stringItem46);
                pivotCacheRecord10.Append(fieldItem39);
                pivotCacheRecord10.Append(stringItem47);
                pivotCacheRecord10.Append(stringItem48);
                pivotCacheRecord10.Append(fieldItem40);

                PivotCacheRecord pivotCacheRecord11 = new PivotCacheRecord();
                FieldItem fieldItem41 = new FieldItem() { Val = (UInt32Value)0U };
                FieldItem fieldItem42 = new FieldItem() { Val = (UInt32Value)0U };
                StringItem stringItem49 = new StringItem() { Val = "231" };
                NumberItem numberItem21 = new NumberItem() { Val = 697D };
                NumberItem numberItem22 = new NumberItem() { Val = 139.4D };
                StringItem stringItem50 = new StringItem() { Val = "0718" };
                FieldItem fieldItem43 = new FieldItem() { Val = (UInt32Value)0U };
                StringItem stringItem51 = new StringItem() { Val = "Агентський договір" };
                StringItem stringItem52 = new StringItem() { Val = "501898" };
                FieldItem fieldItem44 = new FieldItem() { Val = (UInt32Value)0U };

                pivotCacheRecord11.Append(fieldItem41);
                pivotCacheRecord11.Append(fieldItem42);
                pivotCacheRecord11.Append(stringItem49);
                pivotCacheRecord11.Append(numberItem21);
                pivotCacheRecord11.Append(numberItem22);
                pivotCacheRecord11.Append(stringItem50);
                pivotCacheRecord11.Append(fieldItem43);
                pivotCacheRecord11.Append(stringItem51);
                pivotCacheRecord11.Append(stringItem52);
                pivotCacheRecord11.Append(fieldItem44);

                PivotCacheRecord pivotCacheRecord12 = new PivotCacheRecord();
                FieldItem fieldItem45 = new FieldItem() { Val = (UInt32Value)0U };
                FieldItem fieldItem46 = new FieldItem() { Val = (UInt32Value)0U };
                StringItem stringItem53 = new StringItem() { Val = "231" };
                NumberItem numberItem23 = new NumberItem() { Val = 3068D };
                NumberItem numberItem24 = new NumberItem() { Val = 613.6D };
                StringItem stringItem54 = new StringItem() { Val = "0790" };
                FieldItem fieldItem47 = new FieldItem() { Val = (UInt32Value)0U };
                StringItem stringItem55 = new StringItem() { Val = "Агентський договір" };
                StringItem stringItem56 = new StringItem() { Val = "501898" };
                FieldItem fieldItem48 = new FieldItem() { Val = (UInt32Value)0U };

                pivotCacheRecord12.Append(fieldItem45);
                pivotCacheRecord12.Append(fieldItem46);
                pivotCacheRecord12.Append(stringItem53);
                pivotCacheRecord12.Append(numberItem23);
                pivotCacheRecord12.Append(numberItem24);
                pivotCacheRecord12.Append(stringItem54);
                pivotCacheRecord12.Append(fieldItem47);
                pivotCacheRecord12.Append(stringItem55);
                pivotCacheRecord12.Append(stringItem56);
                pivotCacheRecord12.Append(fieldItem48);

                PivotCacheRecord pivotCacheRecord13 = new PivotCacheRecord();
                FieldItem fieldItem49 = new FieldItem() { Val = (UInt32Value)0U };
                FieldItem fieldItem50 = new FieldItem() { Val = (UInt32Value)0U };
                StringItem stringItem57 = new StringItem() { Val = "232" };
                NumberItem numberItem25 = new NumberItem() { Val = 3857D };
                NumberItem numberItem26 = new NumberItem() { Val = 771.4D };
                StringItem stringItem58 = new StringItem() { Val = "0790" };
                FieldItem fieldItem51 = new FieldItem() { Val = (UInt32Value)0U };
                StringItem stringItem59 = new StringItem() { Val = "Агентський договір" };
                StringItem stringItem60 = new StringItem() { Val = "501898" };
                FieldItem fieldItem52 = new FieldItem() { Val = (UInt32Value)0U };

                pivotCacheRecord13.Append(fieldItem49);
                pivotCacheRecord13.Append(fieldItem50);
                pivotCacheRecord13.Append(stringItem57);
                pivotCacheRecord13.Append(numberItem25);
                pivotCacheRecord13.Append(numberItem26);
                pivotCacheRecord13.Append(stringItem58);
                pivotCacheRecord13.Append(fieldItem51);
                pivotCacheRecord13.Append(stringItem59);
                pivotCacheRecord13.Append(stringItem60);
                pivotCacheRecord13.Append(fieldItem52);

                pivotCacheRecords1.Append(pivotCacheRecord1);
                pivotCacheRecords1.Append(pivotCacheRecord2);
                pivotCacheRecords1.Append(pivotCacheRecord3);
                pivotCacheRecords1.Append(pivotCacheRecord4);
                pivotCacheRecords1.Append(pivotCacheRecord5);
                pivotCacheRecords1.Append(pivotCacheRecord6);
                pivotCacheRecords1.Append(pivotCacheRecord7);
                pivotCacheRecords1.Append(pivotCacheRecord8);
                pivotCacheRecords1.Append(pivotCacheRecord9);
                pivotCacheRecords1.Append(pivotCacheRecord10);
                pivotCacheRecords1.Append(pivotCacheRecord11);
                pivotCacheRecords1.Append(pivotCacheRecord12);
                pivotCacheRecords1.Append(pivotCacheRecord13);

                pivotTableCacheRecordsPart1.PivotCacheRecords = pivotCacheRecords1;
            }

            // Generates content of calculationChainPart1.
            private void GenerateCalculationChainPart1Content(CalculationChainPart calculationChainPart1)
            {
                CalculationChain calculationChain1 = new CalculationChain();
                CalculationCell calculationCell1 = new CalculationCell() { CellReference = "J2", SheetId = 1, NewLevel = true };
                CalculationCell calculationCell2 = new CalculationCell() { CellReference = "J3", SheetId = 1 };
                CalculationCell calculationCell3 = new CalculationCell() { CellReference = "J4", SheetId = 1 };
                CalculationCell calculationCell4 = new CalculationCell() { CellReference = "J5", SheetId = 1 };
                CalculationCell calculationCell5 = new CalculationCell() { CellReference = "J6", SheetId = 1 };
                CalculationCell calculationCell6 = new CalculationCell() { CellReference = "J7", SheetId = 1 };
                CalculationCell calculationCell7 = new CalculationCell() { CellReference = "J8", SheetId = 1 };
                CalculationCell calculationCell8 = new CalculationCell() { CellReference = "J9", SheetId = 1 };
                CalculationCell calculationCell9 = new CalculationCell() { CellReference = "J10", SheetId = 1 };
                CalculationCell calculationCell10 = new CalculationCell() { CellReference = "J11", SheetId = 1 };
                CalculationCell calculationCell11 = new CalculationCell() { CellReference = "J12", SheetId = 1 };
                CalculationCell calculationCell12 = new CalculationCell() { CellReference = "J13", SheetId = 1 };
                CalculationCell calculationCell13 = new CalculationCell() { CellReference = "J14", SheetId = 1 };

                calculationChain1.Append(calculationCell1);
                calculationChain1.Append(calculationCell2);
                calculationChain1.Append(calculationCell3);
                calculationChain1.Append(calculationCell4);
                calculationChain1.Append(calculationCell5);
                calculationChain1.Append(calculationCell6);
                calculationChain1.Append(calculationCell7);
                calculationChain1.Append(calculationCell8);
                calculationChain1.Append(calculationCell9);
                calculationChain1.Append(calculationCell10);
                calculationChain1.Append(calculationCell11);
                calculationChain1.Append(calculationCell12);
                calculationChain1.Append(calculationCell13);

                calculationChainPart1.CalculationChain = calculationChain1;
            }

            // Generates content of worksheetPart1.
            private void GenerateWorksheetPart1Content(WorksheetPart worksheetPart1)
            {
                Worksheet worksheet1 = new Worksheet() { MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "x14ac" } };
                worksheet1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
                worksheet1.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
                worksheet1.AddNamespaceDeclaration("x14ac", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac");
                SheetDimension sheetDimension1 = new SheetDimension() { Reference = "A1:J14" };

                SheetViews sheetViews1 = new SheetViews();

                SheetView sheetView1 = new SheetView() { TabSelected = true, WorkbookViewId = (UInt32Value)0U };
                Selection selection1 = new Selection() { ActiveCell = "A2", SequenceOfReferences = new ListValue<StringValue>() { InnerText = "A2" } };

                sheetView1.Append(selection1);

                sheetViews1.Append(sheetView1);
                SheetFormatProperties sheetFormatProperties1 = new SheetFormatProperties() { DefaultRowHeight = 15D, DyDescent = 0.25D };

                Columns columns1 = new Columns();
                Column column1 = new Column() { Min = (UInt32Value)1U, Max = (UInt32Value)1U, Width = 11D, BestFit = true, CustomWidth = true };
                Column column2 = new Column() { Min = (UInt32Value)2U, Max = (UInt32Value)2U, Width = 16.140625D, BestFit = true, CustomWidth = true };
                Column column3 = new Column() { Min = (UInt32Value)3U, Max = (UInt32Value)3U, Width = 16.42578125D, BestFit = true, CustomWidth = true };
                Column column4 = new Column() { Min = (UInt32Value)6U, Max = (UInt32Value)6U, Width = 17.42578125D, CustomWidth = true };
                Column column5 = new Column() { Min = (UInt32Value)7U, Max = (UInt32Value)7U, Width = 66D, BestFit = true, CustomWidth = true };
                Column column6 = new Column() { Min = (UInt32Value)8U, Max = (UInt32Value)8U, Width = 18.85546875D, BestFit = true, CustomWidth = true };
                Column column7 = new Column() { Min = (UInt32Value)9U, Max = (UInt32Value)9U, Width = 9.28515625D, CustomWidth = true };
                Column column8 = new Column() { Min = (UInt32Value)10U, Max = (UInt32Value)10U, Width = 11.85546875D, BestFit = true, CustomWidth = true };

                columns1.Append(column1);
                columns1.Append(column2);
                columns1.Append(column3);
                columns1.Append(column4);
                columns1.Append(column5);
                columns1.Append(column6);
                columns1.Append(column7);
                columns1.Append(column8);

                SheetData sheetData1 = new SheetData();

                Row row1 = new Row() { RowIndex = (UInt32Value)1U, Spans = new ListValue<StringValue>() { InnerText = "1:10" }, DyDescent = 0.25D };

                Cell cell1 = new Cell() { CellReference = "A1", DataType = CellValues.SharedString };
                CellValue cellValue1 = new CellValue();
                cellValue1.Text = "24";

                cell1.Append(cellValue1);

                Cell cell2 = new Cell() { CellReference = "B1", DataType = CellValues.SharedString };
                CellValue cellValue2 = new CellValue();
                cellValue2.Text = "25";

                cell2.Append(cellValue2);

                Cell cell3 = new Cell() { CellReference = "C1", DataType = CellValues.SharedString };
                CellValue cellValue3 = new CellValue();
                cellValue3.Text = "26";

                cell3.Append(cellValue3);

                Cell cell4 = new Cell() { CellReference = "D1", DataType = CellValues.SharedString };
                CellValue cellValue4 = new CellValue();
                cellValue4.Text = "27";

                cell4.Append(cellValue4);

                Cell cell5 = new Cell() { CellReference = "E1", DataType = CellValues.SharedString };
                CellValue cellValue5 = new CellValue();
                cellValue5.Text = "28";

                cell5.Append(cellValue5);

                Cell cell6 = new Cell() { CellReference = "F1", DataType = CellValues.SharedString };
                CellValue cellValue6 = new CellValue();
                cellValue6.Text = "29";

                cell6.Append(cellValue6);

                Cell cell7 = new Cell() { CellReference = "G1", DataType = CellValues.SharedString };
                CellValue cellValue7 = new CellValue();
                cellValue7.Text = "30";

                cell7.Append(cellValue7);

                Cell cell8 = new Cell() { CellReference = "H1", DataType = CellValues.SharedString };
                CellValue cellValue8 = new CellValue();
                cellValue8.Text = "31";

                cell8.Append(cellValue8);

                Cell cell9 = new Cell() { CellReference = "I1", DataType = CellValues.SharedString };
                CellValue cellValue9 = new CellValue();
                cellValue9.Text = "32";

                cell9.Append(cellValue9);

                Cell cell10 = new Cell() { CellReference = "J1", DataType = CellValues.SharedString };
                CellValue cellValue10 = new CellValue();
                cellValue10.Text = "33";

                cell10.Append(cellValue10);

                row1.Append(cell1);
                row1.Append(cell2);
                row1.Append(cell3);
                row1.Append(cell4);
                row1.Append(cell5);
                row1.Append(cell6);
                row1.Append(cell7);
                row1.Append(cell8);
                row1.Append(cell9);
                row1.Append(cell10);

                Row row2 = new Row() { RowIndex = (UInt32Value)2U, Spans = new ListValue<StringValue>() { InnerText = "1:10" }, DyDescent = 0.25D };

                Cell cell11 = new Cell() { CellReference = "A2", StyleIndex = (UInt32Value)1U, DataType = CellValues.SharedString };
                CellValue cellValue11 = new CellValue();
                cellValue11.Text = "0";

                cell11.Append(cellValue11);

                Cell cell12 = new Cell() { CellReference = "B2", DataType = CellValues.SharedString };
                CellValue cellValue12 = new CellValue();
                cellValue12.Text = "1";

                cell12.Append(cellValue12);

                Cell cell13 = new Cell() { CellReference = "C2", StyleIndex = (UInt32Value)1U, DataType = CellValues.SharedString };
                CellValue cellValue13 = new CellValue();
                cellValue13.Text = "2";

                cell13.Append(cellValue13);

                Cell cell14 = new Cell() { CellReference = "D2" };
                CellValue cellValue14 = new CellValue();
                cellValue14.Text = "2917";

                cell14.Append(cellValue14);

                Cell cell15 = new Cell() { CellReference = "E2" };
                CellValue cellValue15 = new CellValue();
                cellValue15.Text = "583.4";

                cell15.Append(cellValue15);

                Cell cell16 = new Cell() { CellReference = "F2", StyleIndex = (UInt32Value)1U, DataType = CellValues.SharedString };
                CellValue cellValue16 = new CellValue();
                cellValue16.Text = "3";

                cell16.Append(cellValue16);

                Cell cell17 = new Cell() { CellReference = "G2", DataType = CellValues.SharedString };
                CellValue cellValue17 = new CellValue();
                cellValue17.Text = "4";

                cell17.Append(cellValue17);

                Cell cell18 = new Cell() { CellReference = "H2", DataType = CellValues.SharedString };
                CellValue cellValue18 = new CellValue();
                cellValue18.Text = "5";

                cell18.Append(cellValue18);

                Cell cell19 = new Cell() { CellReference = "I2", StyleIndex = (UInt32Value)1U, DataType = CellValues.SharedString };
                CellValue cellValue19 = new CellValue();
                cellValue19.Text = "6";

                cell19.Append(cellValue19);

                Cell cell20 = new Cell() { CellReference = "J2", DataType = CellValues.String };
                CellFormula cellFormula1 = new CellFormula();
                cellFormula1.Text = "LEFT(Таблица1[Код відділення],2)";
                CellValue cellValue20 = new CellValue();
                cellValue20.Text = "07";

                cell20.Append(cellFormula1);
                cell20.Append(cellValue20);

                row2.Append(cell11);
                row2.Append(cell12);
                row2.Append(cell13);
                row2.Append(cell14);
                row2.Append(cell15);
                row2.Append(cell16);
                row2.Append(cell17);
                row2.Append(cell18);
                row2.Append(cell19);
                row2.Append(cell20);

                Row row3 = new Row() { RowIndex = (UInt32Value)3U, Spans = new ListValue<StringValue>() { InnerText = "1:10" }, DyDescent = 0.25D };

                Cell cell21 = new Cell() { CellReference = "A3", StyleIndex = (UInt32Value)1U, DataType = CellValues.SharedString };
                CellValue cellValue21 = new CellValue();
                cellValue21.Text = "0";

                cell21.Append(cellValue21);

                Cell cell22 = new Cell() { CellReference = "B3", DataType = CellValues.SharedString };
                CellValue cellValue22 = new CellValue();
                cellValue22.Text = "1";

                cell22.Append(cellValue22);

                Cell cell23 = new Cell() { CellReference = "C3", StyleIndex = (UInt32Value)1U, DataType = CellValues.SharedString };
                CellValue cellValue23 = new CellValue();
                cellValue23.Text = "7";

                cell23.Append(cellValue23);

                Cell cell24 = new Cell() { CellReference = "D3" };
                CellValue cellValue24 = new CellValue();
                cellValue24.Text = "2517";

                cell24.Append(cellValue24);

                Cell cell25 = new Cell() { CellReference = "E3" };
                CellValue cellValue25 = new CellValue();
                cellValue25.Text = "503.4";

                cell25.Append(cellValue25);

                Cell cell26 = new Cell() { CellReference = "F3", StyleIndex = (UInt32Value)1U, DataType = CellValues.SharedString };
                CellValue cellValue26 = new CellValue();
                cellValue26.Text = "3";

                cell26.Append(cellValue26);

                Cell cell27 = new Cell() { CellReference = "G3", DataType = CellValues.SharedString };
                CellValue cellValue27 = new CellValue();
                cellValue27.Text = "4";

                cell27.Append(cellValue27);

                Cell cell28 = new Cell() { CellReference = "H3", DataType = CellValues.SharedString };
                CellValue cellValue28 = new CellValue();
                cellValue28.Text = "5";

                cell28.Append(cellValue28);

                Cell cell29 = new Cell() { CellReference = "I3", StyleIndex = (UInt32Value)1U, DataType = CellValues.SharedString };
                CellValue cellValue29 = new CellValue();
                cellValue29.Text = "6";

                cell29.Append(cellValue29);

                Cell cell30 = new Cell() { CellReference = "J3", DataType = CellValues.String };
                CellFormula cellFormula2 = new CellFormula();
                cellFormula2.Text = "LEFT(Таблица1[Код відділення],2)";
                CellValue cellValue30 = new CellValue();
                cellValue30.Text = "07";

                cell30.Append(cellFormula2);
                cell30.Append(cellValue30);

                row3.Append(cell21);
                row3.Append(cell22);
                row3.Append(cell23);
                row3.Append(cell24);
                row3.Append(cell25);
                row3.Append(cell26);
                row3.Append(cell27);
                row3.Append(cell28);
                row3.Append(cell29);
                row3.Append(cell30);

                Row row4 = new Row() { RowIndex = (UInt32Value)4U, Spans = new ListValue<StringValue>() { InnerText = "1:10" }, DyDescent = 0.25D };

                Cell cell31 = new Cell() { CellReference = "A4", StyleIndex = (UInt32Value)1U, DataType = CellValues.SharedString };
                CellValue cellValue31 = new CellValue();
                cellValue31.Text = "0";

                cell31.Append(cellValue31);

                Cell cell32 = new Cell() { CellReference = "B4", DataType = CellValues.SharedString };
                CellValue cellValue32 = new CellValue();
                cellValue32.Text = "1";

                cell32.Append(cellValue32);

                Cell cell33 = new Cell() { CellReference = "C4", StyleIndex = (UInt32Value)1U, DataType = CellValues.SharedString };
                CellValue cellValue33 = new CellValue();
                cellValue33.Text = "2";

                cell33.Append(cellValue33);

                Cell cell34 = new Cell() { CellReference = "D4" };
                CellValue cellValue34 = new CellValue();
                cellValue34.Text = "2091";

                cell34.Append(cellValue34);

                Cell cell35 = new Cell() { CellReference = "E4" };
                CellValue cellValue35 = new CellValue();
                cellValue35.Text = "418.2";

                cell35.Append(cellValue35);

                Cell cell36 = new Cell() { CellReference = "F4", StyleIndex = (UInt32Value)1U, DataType = CellValues.SharedString };
                CellValue cellValue36 = new CellValue();
                cellValue36.Text = "8";

                cell36.Append(cellValue36);

                Cell cell37 = new Cell() { CellReference = "G4", DataType = CellValues.SharedString };
                CellValue cellValue37 = new CellValue();
                cellValue37.Text = "4";

                cell37.Append(cellValue37);

                Cell cell38 = new Cell() { CellReference = "H4", DataType = CellValues.SharedString };
                CellValue cellValue38 = new CellValue();
                cellValue38.Text = "5";

                cell38.Append(cellValue38);

                Cell cell39 = new Cell() { CellReference = "I4", StyleIndex = (UInt32Value)1U, DataType = CellValues.SharedString };
                CellValue cellValue39 = new CellValue();
                cellValue39.Text = "6";

                cell39.Append(cellValue39);

                Cell cell40 = new Cell() { CellReference = "J4", DataType = CellValues.String };
                CellFormula cellFormula3 = new CellFormula();
                cellFormula3.Text = "LEFT(Таблица1[Код відділення],2)";
                CellValue cellValue40 = new CellValue();
                cellValue40.Text = "07";

                cell40.Append(cellFormula3);
                cell40.Append(cellValue40);

                row4.Append(cell31);
                row4.Append(cell32);
                row4.Append(cell33);
                row4.Append(cell34);
                row4.Append(cell35);
                row4.Append(cell36);
                row4.Append(cell37);
                row4.Append(cell38);
                row4.Append(cell39);
                row4.Append(cell40);

                Row row5 = new Row() { RowIndex = (UInt32Value)5U, Spans = new ListValue<StringValue>() { InnerText = "1:10" }, DyDescent = 0.25D };

                Cell cell41 = new Cell() { CellReference = "A5", StyleIndex = (UInt32Value)1U, DataType = CellValues.SharedString };
                CellValue cellValue41 = new CellValue();
                cellValue41.Text = "0";

                cell41.Append(cellValue41);

                Cell cell42 = new Cell() { CellReference = "B5", DataType = CellValues.SharedString };
                CellValue cellValue42 = new CellValue();
                cellValue42.Text = "1";

                cell42.Append(cellValue42);

                Cell cell43 = new Cell() { CellReference = "C5", StyleIndex = (UInt32Value)1U, DataType = CellValues.SharedString };
                CellValue cellValue43 = new CellValue();
                cellValue43.Text = "2";

                cell43.Append(cellValue43);

                Cell cell44 = new Cell() { CellReference = "D5" };
                CellValue cellValue44 = new CellValue();
                cellValue44.Text = "697";

                cell44.Append(cellValue44);

                Cell cell45 = new Cell() { CellReference = "E5" };
                CellValue cellValue45 = new CellValue();
                cellValue45.Text = "139.4";

                cell45.Append(cellValue45);

                Cell cell46 = new Cell() { CellReference = "F5", StyleIndex = (UInt32Value)1U, DataType = CellValues.SharedString };
                CellValue cellValue46 = new CellValue();
                cellValue46.Text = "9";

                cell46.Append(cellValue46);

                Cell cell47 = new Cell() { CellReference = "G5", DataType = CellValues.SharedString };
                CellValue cellValue47 = new CellValue();
                cellValue47.Text = "4";

                cell47.Append(cellValue47);

                Cell cell48 = new Cell() { CellReference = "H5", DataType = CellValues.SharedString };
                CellValue cellValue48 = new CellValue();
                cellValue48.Text = "5";

                cell48.Append(cellValue48);

                Cell cell49 = new Cell() { CellReference = "I5", StyleIndex = (UInt32Value)1U, DataType = CellValues.SharedString };
                CellValue cellValue49 = new CellValue();
                cellValue49.Text = "6";

                cell49.Append(cellValue49);

                Cell cell50 = new Cell() { CellReference = "J5", DataType = CellValues.String };
                CellFormula cellFormula4 = new CellFormula();
                cellFormula4.Text = "LEFT(Таблица1[Код відділення],2)";
                CellValue cellValue50 = new CellValue();
                cellValue50.Text = "07";

                cell50.Append(cellFormula4);
                cell50.Append(cellValue50);

                row5.Append(cell41);
                row5.Append(cell42);
                row5.Append(cell43);
                row5.Append(cell44);
                row5.Append(cell45);
                row5.Append(cell46);
                row5.Append(cell47);
                row5.Append(cell48);
                row5.Append(cell49);
                row5.Append(cell50);

                Row row6 = new Row() { RowIndex = (UInt32Value)6U, Spans = new ListValue<StringValue>() { InnerText = "1:10" }, DyDescent = 0.25D };

                Cell cell51 = new Cell() { CellReference = "A6", StyleIndex = (UInt32Value)1U, DataType = CellValues.SharedString };
                CellValue cellValue51 = new CellValue();
                cellValue51.Text = "0";

                cell51.Append(cellValue51);

                Cell cell52 = new Cell() { CellReference = "B6", DataType = CellValues.SharedString };
                CellValue cellValue52 = new CellValue();
                cellValue52.Text = "1";

                cell52.Append(cellValue52);

                Cell cell53 = new Cell() { CellReference = "C6", StyleIndex = (UInt32Value)1U, DataType = CellValues.SharedString };
                CellValue cellValue53 = new CellValue();
                cellValue53.Text = "2";

                cell53.Append(cellValue53);

                Cell cell54 = new Cell() { CellReference = "D6" };
                CellValue cellValue54 = new CellValue();
                cellValue54.Text = "6102";

                cell54.Append(cellValue54);

                Cell cell55 = new Cell() { CellReference = "E6" };
                CellValue cellValue55 = new CellValue();
                cellValue55.Text = "1220.4000000000001";

                cell55.Append(cellValue55);

                Cell cell56 = new Cell() { CellReference = "F6", StyleIndex = (UInt32Value)1U, DataType = CellValues.SharedString };
                CellValue cellValue56 = new CellValue();
                cellValue56.Text = "10";

                cell56.Append(cellValue56);

                Cell cell57 = new Cell() { CellReference = "G6", DataType = CellValues.SharedString };
                CellValue cellValue57 = new CellValue();
                cellValue57.Text = "4";

                cell57.Append(cellValue57);

                Cell cell58 = new Cell() { CellReference = "H6", DataType = CellValues.SharedString };
                CellValue cellValue58 = new CellValue();
                cellValue58.Text = "5";

                cell58.Append(cellValue58);

                Cell cell59 = new Cell() { CellReference = "I6", StyleIndex = (UInt32Value)1U, DataType = CellValues.SharedString };
                CellValue cellValue59 = new CellValue();
                cellValue59.Text = "6";

                cell59.Append(cellValue59);

                Cell cell60 = new Cell() { CellReference = "J6", DataType = CellValues.String };
                CellFormula cellFormula5 = new CellFormula();
                cellFormula5.Text = "LEFT(Таблица1[Код відділення],2)";
                CellValue cellValue60 = new CellValue();
                cellValue60.Text = "07";

                cell60.Append(cellFormula5);
                cell60.Append(cellValue60);

                row6.Append(cell51);
                row6.Append(cell52);
                row6.Append(cell53);
                row6.Append(cell54);
                row6.Append(cell55);
                row6.Append(cell56);
                row6.Append(cell57);
                row6.Append(cell58);
                row6.Append(cell59);
                row6.Append(cell60);

                Row row7 = new Row() { RowIndex = (UInt32Value)7U, Spans = new ListValue<StringValue>() { InnerText = "1:10" }, DyDescent = 0.25D };

                Cell cell61 = new Cell() { CellReference = "A7", StyleIndex = (UInt32Value)1U, DataType = CellValues.SharedString };
                CellValue cellValue61 = new CellValue();
                cellValue61.Text = "0";

                cell61.Append(cellValue61);

                Cell cell62 = new Cell() { CellReference = "B7", DataType = CellValues.SharedString };
                CellValue cellValue62 = new CellValue();
                cellValue62.Text = "1";

                cell62.Append(cellValue62);

                Cell cell63 = new Cell() { CellReference = "C7", StyleIndex = (UInt32Value)1U, DataType = CellValues.SharedString };
                CellValue cellValue63 = new CellValue();
                cellValue63.Text = "2";

                cell63.Append(cellValue63);

                Cell cell64 = new Cell() { CellReference = "D7" };
                CellValue cellValue64 = new CellValue();
                cellValue64.Text = "697";

                cell64.Append(cellValue64);

                Cell cell65 = new Cell() { CellReference = "E7" };
                CellValue cellValue65 = new CellValue();
                cellValue65.Text = "139.4";

                cell65.Append(cellValue65);

                Cell cell66 = new Cell() { CellReference = "F7", StyleIndex = (UInt32Value)1U, DataType = CellValues.SharedString };
                CellValue cellValue66 = new CellValue();
                cellValue66.Text = "11";

                cell66.Append(cellValue66);

                Cell cell67 = new Cell() { CellReference = "G7", DataType = CellValues.SharedString };
                CellValue cellValue67 = new CellValue();
                cellValue67.Text = "4";

                cell67.Append(cellValue67);

                Cell cell68 = new Cell() { CellReference = "H7", DataType = CellValues.SharedString };
                CellValue cellValue68 = new CellValue();
                cellValue68.Text = "5";

                cell68.Append(cellValue68);

                Cell cell69 = new Cell() { CellReference = "I7", StyleIndex = (UInt32Value)1U, DataType = CellValues.SharedString };
                CellValue cellValue69 = new CellValue();
                cellValue69.Text = "6";

                cell69.Append(cellValue69);

                Cell cell70 = new Cell() { CellReference = "J7", DataType = CellValues.String };
                CellFormula cellFormula6 = new CellFormula();
                cellFormula6.Text = "LEFT(Таблица1[Код відділення],2)";
                CellValue cellValue70 = new CellValue();
                cellValue70.Text = "07";

                cell70.Append(cellFormula6);
                cell70.Append(cellValue70);

                row7.Append(cell61);
                row7.Append(cell62);
                row7.Append(cell63);
                row7.Append(cell64);
                row7.Append(cell65);
                row7.Append(cell66);
                row7.Append(cell67);
                row7.Append(cell68);
                row7.Append(cell69);
                row7.Append(cell70);

                Row row8 = new Row() { RowIndex = (UInt32Value)8U, Spans = new ListValue<StringValue>() { InnerText = "1:10" }, DyDescent = 0.25D };

                Cell cell71 = new Cell() { CellReference = "A8", StyleIndex = (UInt32Value)1U, DataType = CellValues.SharedString };
                CellValue cellValue71 = new CellValue();
                cellValue71.Text = "0";

                cell71.Append(cellValue71);

                Cell cell72 = new Cell() { CellReference = "B8", DataType = CellValues.SharedString };
                CellValue cellValue72 = new CellValue();
                cellValue72.Text = "1";

                cell72.Append(cellValue72);

                Cell cell73 = new Cell() { CellReference = "C8", StyleIndex = (UInt32Value)1U, DataType = CellValues.SharedString };
                CellValue cellValue73 = new CellValue();
                cellValue73.Text = "2";

                cell73.Append(cellValue73);

                Cell cell74 = new Cell() { CellReference = "D8" };
                CellValue cellValue74 = new CellValue();
                cellValue74.Text = "6102";

                cell74.Append(cellValue74);

                Cell cell75 = new Cell() { CellReference = "E8" };
                CellValue cellValue75 = new CellValue();
                cellValue75.Text = "1220.4000000000001";

                cell75.Append(cellValue75);

                Cell cell76 = new Cell() { CellReference = "F8", StyleIndex = (UInt32Value)1U, DataType = CellValues.SharedString };
                CellValue cellValue76 = new CellValue();
                cellValue76.Text = "12";

                cell76.Append(cellValue76);

                Cell cell77 = new Cell() { CellReference = "G8", DataType = CellValues.SharedString };
                CellValue cellValue77 = new CellValue();
                cellValue77.Text = "4";

                cell77.Append(cellValue77);

                Cell cell78 = new Cell() { CellReference = "H8", DataType = CellValues.SharedString };
                CellValue cellValue78 = new CellValue();
                cellValue78.Text = "5";

                cell78.Append(cellValue78);

                Cell cell79 = new Cell() { CellReference = "I8", StyleIndex = (UInt32Value)1U, DataType = CellValues.SharedString };
                CellValue cellValue79 = new CellValue();
                cellValue79.Text = "6";

                cell79.Append(cellValue79);

                Cell cell80 = new Cell() { CellReference = "J8", DataType = CellValues.String };
                CellFormula cellFormula7 = new CellFormula();
                cellFormula7.Text = "LEFT(Таблица1[Код відділення],2)";
                CellValue cellValue80 = new CellValue();
                cellValue80.Text = "07";

                cell80.Append(cellFormula7);
                cell80.Append(cellValue80);

                row8.Append(cell71);
                row8.Append(cell72);
                row8.Append(cell73);
                row8.Append(cell74);
                row8.Append(cell75);
                row8.Append(cell76);
                row8.Append(cell77);
                row8.Append(cell78);
                row8.Append(cell79);
                row8.Append(cell80);

                Row row9 = new Row() { RowIndex = (UInt32Value)9U, Spans = new ListValue<StringValue>() { InnerText = "1:10" }, DyDescent = 0.25D };

                Cell cell81 = new Cell() { CellReference = "A9", StyleIndex = (UInt32Value)1U, DataType = CellValues.SharedString };
                CellValue cellValue81 = new CellValue();
                cellValue81.Text = "13";

                cell81.Append(cellValue81);

                Cell cell82 = new Cell() { CellReference = "B9", DataType = CellValues.SharedString };
                CellValue cellValue82 = new CellValue();
                cellValue82.Text = "14";

                cell82.Append(cellValue82);

                Cell cell83 = new Cell() { CellReference = "C9", StyleIndex = (UInt32Value)1U, DataType = CellValues.SharedString };
                CellValue cellValue83 = new CellValue();
                cellValue83.Text = "2";

                cell83.Append(cellValue83);

                Cell cell84 = new Cell() { CellReference = "D9" };
                CellValue cellValue84 = new CellValue();
                cellValue84.Text = "1807";

                cell84.Append(cellValue84);

                Cell cell85 = new Cell() { CellReference = "E9" };
                CellValue cellValue85 = new CellValue();
                cellValue85.Text = "361.4";

                cell85.Append(cellValue85);

                Cell cell86 = new Cell() { CellReference = "F9", StyleIndex = (UInt32Value)1U, DataType = CellValues.SharedString };
                CellValue cellValue86 = new CellValue();
                cellValue86.Text = "15";

                cell86.Append(cellValue86);

                Cell cell87 = new Cell() { CellReference = "G9", DataType = CellValues.SharedString };
                CellValue cellValue87 = new CellValue();
                cellValue87.Text = "4";

                cell87.Append(cellValue87);

                Cell cell88 = new Cell() { CellReference = "H9", DataType = CellValues.SharedString };
                CellValue cellValue88 = new CellValue();
                cellValue88.Text = "5";

                cell88.Append(cellValue88);

                Cell cell89 = new Cell() { CellReference = "I9", StyleIndex = (UInt32Value)1U, DataType = CellValues.SharedString };
                CellValue cellValue89 = new CellValue();
                cellValue89.Text = "16";

                cell89.Append(cellValue89);

                Cell cell90 = new Cell() { CellReference = "J9", DataType = CellValues.String };
                CellFormula cellFormula8 = new CellFormula();
                cellFormula8.Text = "LEFT(Таблица1[Код відділення],2)";
                CellValue cellValue90 = new CellValue();
                cellValue90.Text = "07";

                cell90.Append(cellFormula8);
                cell90.Append(cellValue90);

                row9.Append(cell81);
                row9.Append(cell82);
                row9.Append(cell83);
                row9.Append(cell84);
                row9.Append(cell85);
                row9.Append(cell86);
                row9.Append(cell87);
                row9.Append(cell88);
                row9.Append(cell89);
                row9.Append(cell90);

                Row row10 = new Row() { RowIndex = (UInt32Value)10U, Spans = new ListValue<StringValue>() { InnerText = "1:10" }, DyDescent = 0.25D };

                Cell cell91 = new Cell() { CellReference = "A10", StyleIndex = (UInt32Value)1U, DataType = CellValues.SharedString };
                CellValue cellValue91 = new CellValue();
                cellValue91.Text = "17";

                cell91.Append(cellValue91);

                Cell cell92 = new Cell() { CellReference = "B10", DataType = CellValues.SharedString };
                CellValue cellValue92 = new CellValue();
                cellValue92.Text = "18";

                cell92.Append(cellValue92);

                Cell cell93 = new Cell() { CellReference = "C10", StyleIndex = (UInt32Value)1U, DataType = CellValues.SharedString };
                CellValue cellValue93 = new CellValue();
                cellValue93.Text = "2";

                cell93.Append(cellValue93);

                Cell cell94 = new Cell() { CellReference = "D10" };
                CellValue cellValue94 = new CellValue();
                cellValue94.Text = "7987";

                cell94.Append(cellValue94);

                Cell cell95 = new Cell() { CellReference = "E10" };
                CellValue cellValue95 = new CellValue();
                cellValue95.Text = "1597.4";

                cell95.Append(cellValue95);

                Cell cell96 = new Cell() { CellReference = "F10", StyleIndex = (UInt32Value)1U, DataType = CellValues.SharedString };
                CellValue cellValue96 = new CellValue();
                cellValue96.Text = "19";

                cell96.Append(cellValue96);

                Cell cell97 = new Cell() { CellReference = "G10", DataType = CellValues.SharedString };
                CellValue cellValue97 = new CellValue();
                cellValue97.Text = "4";

                cell97.Append(cellValue97);

                Cell cell98 = new Cell() { CellReference = "H10", DataType = CellValues.SharedString };
                CellValue cellValue98 = new CellValue();
                cellValue98.Text = "5";

                cell98.Append(cellValue98);

                Cell cell99 = new Cell() { CellReference = "I10", StyleIndex = (UInt32Value)1U, DataType = CellValues.SharedString };
                CellValue cellValue99 = new CellValue();
                cellValue99.Text = "20";

                cell99.Append(cellValue99);

                Cell cell100 = new Cell() { CellReference = "J10", DataType = CellValues.String };
                CellFormula cellFormula9 = new CellFormula();
                cellFormula9.Text = "LEFT(Таблица1[Код відділення],2)";
                CellValue cellValue100 = new CellValue();
                cellValue100.Text = "07";

                cell100.Append(cellFormula9);
                cell100.Append(cellValue100);

                row10.Append(cell91);
                row10.Append(cell92);
                row10.Append(cell93);
                row10.Append(cell94);
                row10.Append(cell95);
                row10.Append(cell96);
                row10.Append(cell97);
                row10.Append(cell98);
                row10.Append(cell99);
                row10.Append(cell100);

                Row row11 = new Row() { RowIndex = (UInt32Value)11U, Spans = new ListValue<StringValue>() { InnerText = "1:10" }, DyDescent = 0.25D };

                Cell cell101 = new Cell() { CellReference = "A11", StyleIndex = (UInt32Value)1U, DataType = CellValues.SharedString };
                CellValue cellValue101 = new CellValue();
                cellValue101.Text = "0";

                cell101.Append(cellValue101);

                Cell cell102 = new Cell() { CellReference = "B11", DataType = CellValues.SharedString };
                CellValue cellValue102 = new CellValue();
                cellValue102.Text = "1";

                cell102.Append(cellValue102);

                Cell cell103 = new Cell() { CellReference = "C11", StyleIndex = (UInt32Value)1U, DataType = CellValues.SharedString };
                CellValue cellValue103 = new CellValue();
                cellValue103.Text = "2";

                cell103.Append(cellValue103);

                Cell cell104 = new Cell() { CellReference = "D11" };
                CellValue cellValue104 = new CellValue();
                cellValue104.Text = "697";

                cell104.Append(cellValue104);

                Cell cell105 = new Cell() { CellReference = "E11" };
                CellValue cellValue105 = new CellValue();
                cellValue105.Text = "139.4";

                cell105.Append(cellValue105);

                Cell cell106 = new Cell() { CellReference = "F11", StyleIndex = (UInt32Value)1U, DataType = CellValues.SharedString };
                CellValue cellValue106 = new CellValue();
                cellValue106.Text = "21";

                cell106.Append(cellValue106);

                Cell cell107 = new Cell() { CellReference = "G11", DataType = CellValues.SharedString };
                CellValue cellValue107 = new CellValue();
                cellValue107.Text = "4";

                cell107.Append(cellValue107);

                Cell cell108 = new Cell() { CellReference = "H11", DataType = CellValues.SharedString };
                CellValue cellValue108 = new CellValue();
                cellValue108.Text = "5";

                cell108.Append(cellValue108);

                Cell cell109 = new Cell() { CellReference = "I11", StyleIndex = (UInt32Value)1U, DataType = CellValues.SharedString };
                CellValue cellValue109 = new CellValue();
                cellValue109.Text = "6";

                cell109.Append(cellValue109);

                Cell cell110 = new Cell() { CellReference = "J11", DataType = CellValues.String };
                CellFormula cellFormula10 = new CellFormula();
                cellFormula10.Text = "LEFT(Таблица1[Код відділення],2)";
                CellValue cellValue110 = new CellValue();
                cellValue110.Text = "07";

                cell110.Append(cellFormula10);
                cell110.Append(cellValue110);

                row11.Append(cell101);
                row11.Append(cell102);
                row11.Append(cell103);
                row11.Append(cell104);
                row11.Append(cell105);
                row11.Append(cell106);
                row11.Append(cell107);
                row11.Append(cell108);
                row11.Append(cell109);
                row11.Append(cell110);

                Row row12 = new Row() { RowIndex = (UInt32Value)12U, Spans = new ListValue<StringValue>() { InnerText = "1:10" }, DyDescent = 0.25D };

                Cell cell111 = new Cell() { CellReference = "A12", StyleIndex = (UInt32Value)1U, DataType = CellValues.SharedString };
                CellValue cellValue111 = new CellValue();
                cellValue111.Text = "0";

                cell111.Append(cellValue111);

                Cell cell112 = new Cell() { CellReference = "B12", DataType = CellValues.SharedString };
                CellValue cellValue112 = new CellValue();
                cellValue112.Text = "1";

                cell112.Append(cellValue112);

                Cell cell113 = new Cell() { CellReference = "C12", StyleIndex = (UInt32Value)1U, DataType = CellValues.SharedString };
                CellValue cellValue113 = new CellValue();
                cellValue113.Text = "2";

                cell113.Append(cellValue113);

                Cell cell114 = new Cell() { CellReference = "D12" };
                CellValue cellValue114 = new CellValue();
                cellValue114.Text = "697";

                cell114.Append(cellValue114);

                Cell cell115 = new Cell() { CellReference = "E12" };
                CellValue cellValue115 = new CellValue();
                cellValue115.Text = "139.4";

                cell115.Append(cellValue115);

                Cell cell116 = new Cell() { CellReference = "F12", StyleIndex = (UInt32Value)1U, DataType = CellValues.SharedString };
                CellValue cellValue116 = new CellValue();
                cellValue116.Text = "22";

                cell116.Append(cellValue116);

                Cell cell117 = new Cell() { CellReference = "G12", DataType = CellValues.SharedString };
                CellValue cellValue117 = new CellValue();
                cellValue117.Text = "4";

                cell117.Append(cellValue117);

                Cell cell118 = new Cell() { CellReference = "H12", DataType = CellValues.SharedString };
                CellValue cellValue118 = new CellValue();
                cellValue118.Text = "5";

                cell118.Append(cellValue118);

                Cell cell119 = new Cell() { CellReference = "I12", StyleIndex = (UInt32Value)1U, DataType = CellValues.SharedString };
                CellValue cellValue119 = new CellValue();
                cellValue119.Text = "6";

                cell119.Append(cellValue119);

                Cell cell120 = new Cell() { CellReference = "J12", DataType = CellValues.String };
                CellFormula cellFormula11 = new CellFormula();
                cellFormula11.Text = "LEFT(Таблица1[Код відділення],2)";
                CellValue cellValue120 = new CellValue();
                cellValue120.Text = "07";

                cell120.Append(cellFormula11);
                cell120.Append(cellValue120);

                row12.Append(cell111);
                row12.Append(cell112);
                row12.Append(cell113);
                row12.Append(cell114);
                row12.Append(cell115);
                row12.Append(cell116);
                row12.Append(cell117);
                row12.Append(cell118);
                row12.Append(cell119);
                row12.Append(cell120);

                Row row13 = new Row() { RowIndex = (UInt32Value)13U, Spans = new ListValue<StringValue>() { InnerText = "1:10" }, DyDescent = 0.25D };

                Cell cell121 = new Cell() { CellReference = "A13", StyleIndex = (UInt32Value)1U, DataType = CellValues.SharedString };
                CellValue cellValue121 = new CellValue();
                cellValue121.Text = "0";

                cell121.Append(cellValue121);

                Cell cell122 = new Cell() { CellReference = "B13", DataType = CellValues.SharedString };
                CellValue cellValue122 = new CellValue();
                cellValue122.Text = "1";

                cell122.Append(cellValue122);

                Cell cell123 = new Cell() { CellReference = "C13", StyleIndex = (UInt32Value)1U, DataType = CellValues.SharedString };
                CellValue cellValue123 = new CellValue();
                cellValue123.Text = "2";

                cell123.Append(cellValue123);

                Cell cell124 = new Cell() { CellReference = "D13" };
                CellValue cellValue124 = new CellValue();
                cellValue124.Text = "3068";

                cell124.Append(cellValue124);

                Cell cell125 = new Cell() { CellReference = "E13" };
                CellValue cellValue125 = new CellValue();
                cellValue125.Text = "613.6";

                cell125.Append(cellValue125);

                Cell cell126 = new Cell() { CellReference = "F13", StyleIndex = (UInt32Value)1U, DataType = CellValues.SharedString };
                CellValue cellValue126 = new CellValue();
                cellValue126.Text = "23";

                cell126.Append(cellValue126);

                Cell cell127 = new Cell() { CellReference = "G13", DataType = CellValues.SharedString };
                CellValue cellValue127 = new CellValue();
                cellValue127.Text = "4";

                cell127.Append(cellValue127);

                Cell cell128 = new Cell() { CellReference = "H13", DataType = CellValues.SharedString };
                CellValue cellValue128 = new CellValue();
                cellValue128.Text = "5";

                cell128.Append(cellValue128);

                Cell cell129 = new Cell() { CellReference = "I13", StyleIndex = (UInt32Value)1U, DataType = CellValues.SharedString };
                CellValue cellValue129 = new CellValue();
                cellValue129.Text = "6";

                cell129.Append(cellValue129);

                Cell cell130 = new Cell() { CellReference = "J13", DataType = CellValues.String };
                CellFormula cellFormula12 = new CellFormula();
                cellFormula12.Text = "LEFT(Таблица1[Код відділення],2)";
                CellValue cellValue130 = new CellValue();
                cellValue130.Text = "07";

                cell130.Append(cellFormula12);
                cell130.Append(cellValue130);

                row13.Append(cell121);
                row13.Append(cell122);
                row13.Append(cell123);
                row13.Append(cell124);
                row13.Append(cell125);
                row13.Append(cell126);
                row13.Append(cell127);
                row13.Append(cell128);
                row13.Append(cell129);
                row13.Append(cell130);

                Row row14 = new Row() { RowIndex = (UInt32Value)14U, Spans = new ListValue<StringValue>() { InnerText = "1:10" }, DyDescent = 0.25D };

                Cell cell131 = new Cell() { CellReference = "A14", StyleIndex = (UInt32Value)1U, DataType = CellValues.SharedString };
                CellValue cellValue131 = new CellValue();
                cellValue131.Text = "0";

                cell131.Append(cellValue131);

                Cell cell132 = new Cell() { CellReference = "B14", DataType = CellValues.SharedString };
                CellValue cellValue132 = new CellValue();
                cellValue132.Text = "1";

                cell132.Append(cellValue132);

                Cell cell133 = new Cell() { CellReference = "C14", StyleIndex = (UInt32Value)1U, DataType = CellValues.SharedString };
                CellValue cellValue133 = new CellValue();
                cellValue133.Text = "7";

                cell133.Append(cellValue133);

                Cell cell134 = new Cell() { CellReference = "D14" };
                CellValue cellValue134 = new CellValue();
                cellValue134.Text = "3857";

                cell134.Append(cellValue134);

                Cell cell135 = new Cell() { CellReference = "E14" };
                CellValue cellValue135 = new CellValue();
                cellValue135.Text = "771.4";

                cell135.Append(cellValue135);

                Cell cell136 = new Cell() { CellReference = "F14", StyleIndex = (UInt32Value)1U, DataType = CellValues.SharedString };
                CellValue cellValue136 = new CellValue();
                cellValue136.Text = "23";

                cell136.Append(cellValue136);

                Cell cell137 = new Cell() { CellReference = "G14", DataType = CellValues.SharedString };
                CellValue cellValue137 = new CellValue();
                cellValue137.Text = "4";

                cell137.Append(cellValue137);

                Cell cell138 = new Cell() { CellReference = "H14", DataType = CellValues.SharedString };
                CellValue cellValue138 = new CellValue();
                cellValue138.Text = "5";

                cell138.Append(cellValue138);

                Cell cell139 = new Cell() { CellReference = "I14", StyleIndex = (UInt32Value)1U, DataType = CellValues.SharedString };
                CellValue cellValue139 = new CellValue();
                cellValue139.Text = "6";

                cell139.Append(cellValue139);

                Cell cell140 = new Cell() { CellReference = "J14", DataType = CellValues.String };
                CellFormula cellFormula13 = new CellFormula();
                cellFormula13.Text = "LEFT(Таблица1[Код відділення],2)";
                CellValue cellValue140 = new CellValue();
                cellValue140.Text = "07";

                cell140.Append(cellFormula13);
                cell140.Append(cellValue140);

                row14.Append(cell131);
                row14.Append(cell132);
                row14.Append(cell133);
                row14.Append(cell134);
                row14.Append(cell135);
                row14.Append(cell136);
                row14.Append(cell137);
                row14.Append(cell138);
                row14.Append(cell139);
                row14.Append(cell140);

                sheetData1.Append(row1);
                sheetData1.Append(row2);
                sheetData1.Append(row3);
                sheetData1.Append(row4);
                sheetData1.Append(row5);
                sheetData1.Append(row6);
                sheetData1.Append(row7);
                sheetData1.Append(row8);
                sheetData1.Append(row9);
                sheetData1.Append(row10);
                sheetData1.Append(row11);
                sheetData1.Append(row12);
                sheetData1.Append(row13);
                sheetData1.Append(row14);
                PageMargins pageMargins1 = new PageMargins() { Left = 0.7D, Right = 0.7D, Top = 0.75D, Bottom = 0.75D, Header = 0.3D, Footer = 0.3D };

                TableParts tableParts1 = new TableParts() { Count = (UInt32Value)1U };
                TablePart tablePart1 = new TablePart() { Id = "rId1" };

                tableParts1.Append(tablePart1);

                worksheet1.Append(sheetDimension1);
                worksheet1.Append(sheetViews1);
                worksheet1.Append(sheetFormatProperties1);
                worksheet1.Append(columns1);
                worksheet1.Append(sheetData1);
                worksheet1.Append(pageMargins1);
                worksheet1.Append(tableParts1);

                worksheetPart1.Worksheet = worksheet1;
            }

            // Generates content of tableDefinitionPart1.
            private void GenerateTableDefinitionPart1Content(TableDefinitionPart tableDefinitionPart1)
            {
                Table table1 = new Table() { Id = (UInt32Value)1U, Name = "Таблица1", DisplayName = "Таблица1", Reference = "A1:J14", TotalsRowShown = false };
                AutoFilter autoFilter1 = new AutoFilter() { Reference = "A1:J14" };

                TableColumns tableColumns1 = new TableColumns() { Count = (UInt32Value)10U };
                TableColumn tableColumn1 = new TableColumn() { Id = (UInt32Value)1U, Name = "ІПН", DataFormatId = (UInt32Value)4U };
                TableColumn tableColumn2 = new TableColumn() { Id = (UInt32Value)2U, Name = "Агент" };
                TableColumn tableColumn3 = new TableColumn() { Id = (UInt32Value)3U, Name = "Код програми", DataFormatId = (UInt32Value)3U };
                TableColumn tableColumn4 = new TableColumn() { Id = (UInt32Value)4U, Name = "СП" };
                TableColumn tableColumn5 = new TableColumn() { Id = (UInt32Value)5U, Name = "АВ" };
                TableColumn tableColumn6 = new TableColumn() { Id = (UInt32Value)6U, Name = "Код відділення", DataFormatId = (UInt32Value)2U };
                TableColumn tableColumn7 = new TableColumn() { Id = (UInt32Value)7U, Name = "Канал" };
                TableColumn tableColumn8 = new TableColumn() { Id = (UInt32Value)8U, Name = "Договір" };
                TableColumn tableColumn9 = new TableColumn() { Id = (UInt32Value)9U, Name = "ID акту", DataFormatId = (UInt32Value)1U };

                TableColumn tableColumn10 = new TableColumn() { Id = (UInt32Value)10U, Name = "Дирекція", DataFormatId = (UInt32Value)0U };
                CalculatedColumnFormula calculatedColumnFormula1 = new CalculatedColumnFormula();
                calculatedColumnFormula1.Text = "LEFT(Таблица1[Код відділення],2)";

                tableColumn10.Append(calculatedColumnFormula1);

                tableColumns1.Append(tableColumn1);
                tableColumns1.Append(tableColumn2);
                tableColumns1.Append(tableColumn3);
                tableColumns1.Append(tableColumn4);
                tableColumns1.Append(tableColumn5);
                tableColumns1.Append(tableColumn6);
                tableColumns1.Append(tableColumn7);
                tableColumns1.Append(tableColumn8);
                tableColumns1.Append(tableColumn9);
                tableColumns1.Append(tableColumn10);
                TableStyleInfo tableStyleInfo1 = new TableStyleInfo() { Name = "TableStyleMedium2", ShowFirstColumn = false, ShowLastColumn = false, ShowRowStripes = true, ShowColumnStripes = false };

                table1.Append(autoFilter1);
                table1.Append(tableColumns1);
                table1.Append(tableStyleInfo1);

                tableDefinitionPart1.Table = table1;
            }

            // Generates content of worksheetPart2.
            private void GenerateWorksheetPart2Content(WorksheetPart worksheetPart2)
            {
                Worksheet worksheet2 = new Worksheet() { MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "x14ac" } };
                worksheet2.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
                worksheet2.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
                worksheet2.AddNamespaceDeclaration("x14ac", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac");
                SheetDimension sheetDimension2 = new SheetDimension() { Reference = "A3:C12" };

                SheetViews sheetViews2 = new SheetViews();
                SheetView sheetView2 = new SheetView() { WorkbookViewId = (UInt32Value)0U };

                sheetViews2.Append(sheetView2);
                SheetFormatProperties sheetFormatProperties2 = new SheetFormatProperties() { DefaultRowHeight = 15D, DyDescent = 0.25D };

                Columns columns2 = new Columns();
                Column column9 = new Column() { Min = (UInt32Value)1U, Max = (UInt32Value)1U, Width = 70.5703125D, BestFit = true, CustomWidth = true };
                Column column10 = new Column() { Min = (UInt32Value)2U, Max = (UInt32Value)3U, Width = 18.5703125D, BestFit = true, CustomWidth = true };

                columns2.Append(column9);
                columns2.Append(column10);

                SheetData sheetData2 = new SheetData();

                Row row15 = new Row() { RowIndex = (UInt32Value)3U, Spans = new ListValue<StringValue>() { InnerText = "1:3" }, DyDescent = 0.25D };

                Cell cell141 = new Cell() { CellReference = "A3", StyleIndex = (UInt32Value)2U, DataType = CellValues.SharedString };
                CellValue cellValue141 = new CellValue();
                cellValue141.Text = "34";

                cell141.Append(cellValue141);

                Cell cell142 = new Cell() { CellReference = "B3", DataType = CellValues.SharedString };
                CellValue cellValue142 = new CellValue();
                cellValue142.Text = "37";

                cell142.Append(cellValue142);

                Cell cell143 = new Cell() { CellReference = "C3", DataType = CellValues.SharedString };
                CellValue cellValue143 = new CellValue();
                cellValue143.Text = "38";

                cell143.Append(cellValue143);

                row15.Append(cell141);
                row15.Append(cell142);
                row15.Append(cell143);

                Row row16 = new Row() { RowIndex = (UInt32Value)4U, Spans = new ListValue<StringValue>() { InnerText = "1:3" }, DyDescent = 0.25D };

                Cell cell144 = new Cell() { CellReference = "A4", StyleIndex = (UInt32Value)3U, DataType = CellValues.SharedString };
                CellValue cellValue144 = new CellValue();
                cellValue144.Text = "35";

                cell144.Append(cellValue144);

                Cell cell145 = new Cell() { CellReference = "B4", StyleIndex = (UInt32Value)7U };
                CellValue cellValue145 = new CellValue();
                cellValue145.Text = "39236";

                cell145.Append(cellValue145);

                Cell cell146 = new Cell() { CellReference = "C4", StyleIndex = (UInt32Value)7U };
                CellValue cellValue146 = new CellValue();
                cellValue146.Text = "7847.1999999999989";

                cell146.Append(cellValue146);

                row16.Append(cell144);
                row16.Append(cell145);
                row16.Append(cell146);

                Row row17 = new Row() { RowIndex = (UInt32Value)5U, Spans = new ListValue<StringValue>() { InnerText = "1:3" }, DyDescent = 0.25D };

                Cell cell147 = new Cell() { CellReference = "A5", StyleIndex = (UInt32Value)4U, DataType = CellValues.SharedString };
                CellValue cellValue147 = new CellValue();
                cellValue147.Text = "4";

                cell147.Append(cellValue147);

                Cell cell148 = new Cell() { CellReference = "B5", StyleIndex = (UInt32Value)7U };
                CellValue cellValue148 = new CellValue();
                cellValue148.Text = "39236";

                cell148.Append(cellValue148);

                Cell cell149 = new Cell() { CellReference = "C5", StyleIndex = (UInt32Value)7U };
                CellValue cellValue149 = new CellValue();
                cellValue149.Text = "7847.1999999999989";

                cell149.Append(cellValue149);

                row17.Append(cell147);
                row17.Append(cell148);
                row17.Append(cell149);

                Row row18 = new Row() { RowIndex = (UInt32Value)6U, Spans = new ListValue<StringValue>() { InnerText = "1:3" }, DyDescent = 0.25D };

                Cell cell150 = new Cell() { CellReference = "A6", StyleIndex = (UInt32Value)5U, DataType = CellValues.SharedString };
                CellValue cellValue150 = new CellValue();
                cellValue150.Text = "18";

                cell150.Append(cellValue150);

                Cell cell151 = new Cell() { CellReference = "B6", StyleIndex = (UInt32Value)7U };
                CellValue cellValue151 = new CellValue();
                cellValue151.Text = "7987";

                cell151.Append(cellValue151);

                Cell cell152 = new Cell() { CellReference = "C6", StyleIndex = (UInt32Value)7U };
                CellValue cellValue152 = new CellValue();
                cellValue152.Text = "1597.4";

                cell152.Append(cellValue152);

                row18.Append(cell150);
                row18.Append(cell151);
                row18.Append(cell152);

                Row row19 = new Row() { RowIndex = (UInt32Value)7U, Spans = new ListValue<StringValue>() { InnerText = "1:3" }, DyDescent = 0.25D };

                Cell cell153 = new Cell() { CellReference = "A7", StyleIndex = (UInt32Value)6U, DataType = CellValues.SharedString };
                CellValue cellValue153 = new CellValue();
                cellValue153.Text = "17";

                cell153.Append(cellValue153);

                Cell cell154 = new Cell() { CellReference = "B7", StyleIndex = (UInt32Value)7U };
                CellValue cellValue154 = new CellValue();
                cellValue154.Text = "7987";

                cell154.Append(cellValue154);

                Cell cell155 = new Cell() { CellReference = "C7", StyleIndex = (UInt32Value)7U };
                CellValue cellValue155 = new CellValue();
                cellValue155.Text = "1597.4";

                cell155.Append(cellValue155);

                row19.Append(cell153);
                row19.Append(cell154);
                row19.Append(cell155);

                Row row20 = new Row() { RowIndex = (UInt32Value)8U, Spans = new ListValue<StringValue>() { InnerText = "1:3" }, DyDescent = 0.25D };

                Cell cell156 = new Cell() { CellReference = "A8", StyleIndex = (UInt32Value)5U, DataType = CellValues.SharedString };
                CellValue cellValue156 = new CellValue();
                cellValue156.Text = "1";

                cell156.Append(cellValue156);

                Cell cell157 = new Cell() { CellReference = "B8", StyleIndex = (UInt32Value)7U };
                CellValue cellValue157 = new CellValue();
                cellValue157.Text = "29442";

                cell157.Append(cellValue157);

                Cell cell158 = new Cell() { CellReference = "C8", StyleIndex = (UInt32Value)7U };
                CellValue cellValue158 = new CellValue();
                cellValue158.Text = "5888.4";

                cell158.Append(cellValue158);

                row20.Append(cell156);
                row20.Append(cell157);
                row20.Append(cell158);

                Row row21 = new Row() { RowIndex = (UInt32Value)9U, Spans = new ListValue<StringValue>() { InnerText = "1:3" }, DyDescent = 0.25D };

                Cell cell159 = new Cell() { CellReference = "A9", StyleIndex = (UInt32Value)6U, DataType = CellValues.SharedString };
                CellValue cellValue159 = new CellValue();
                cellValue159.Text = "0";

                cell159.Append(cellValue159);

                Cell cell160 = new Cell() { CellReference = "B9", StyleIndex = (UInt32Value)7U };
                CellValue cellValue160 = new CellValue();
                cellValue160.Text = "29442";

                cell160.Append(cellValue160);

                Cell cell161 = new Cell() { CellReference = "C9", StyleIndex = (UInt32Value)7U };
                CellValue cellValue161 = new CellValue();
                cellValue161.Text = "5888.4";

                cell161.Append(cellValue161);

                row21.Append(cell159);
                row21.Append(cell160);
                row21.Append(cell161);

                Row row22 = new Row() { RowIndex = (UInt32Value)10U, Spans = new ListValue<StringValue>() { InnerText = "1:3" }, DyDescent = 0.25D };

                Cell cell162 = new Cell() { CellReference = "A10", StyleIndex = (UInt32Value)5U, DataType = CellValues.SharedString };
                CellValue cellValue162 = new CellValue();
                cellValue162.Text = "14";

                cell162.Append(cellValue162);

                Cell cell163 = new Cell() { CellReference = "B10", StyleIndex = (UInt32Value)7U };
                CellValue cellValue163 = new CellValue();
                cellValue163.Text = "1807";

                cell163.Append(cellValue163);

                Cell cell164 = new Cell() { CellReference = "C10", StyleIndex = (UInt32Value)7U };
                CellValue cellValue164 = new CellValue();
                cellValue164.Text = "361.4";

                cell164.Append(cellValue164);

                row22.Append(cell162);
                row22.Append(cell163);
                row22.Append(cell164);

                Row row23 = new Row() { RowIndex = (UInt32Value)11U, Spans = new ListValue<StringValue>() { InnerText = "1:3" }, DyDescent = 0.25D };

                Cell cell165 = new Cell() { CellReference = "A11", StyleIndex = (UInt32Value)6U, DataType = CellValues.SharedString };
                CellValue cellValue165 = new CellValue();
                cellValue165.Text = "13";

                cell165.Append(cellValue165);

                Cell cell166 = new Cell() { CellReference = "B11", StyleIndex = (UInt32Value)7U };
                CellValue cellValue166 = new CellValue();
                cellValue166.Text = "1807";

                cell166.Append(cellValue166);

                Cell cell167 = new Cell() { CellReference = "C11", StyleIndex = (UInt32Value)7U };
                CellValue cellValue167 = new CellValue();
                cellValue167.Text = "361.4";

                cell167.Append(cellValue167);

                row23.Append(cell165);
                row23.Append(cell166);
                row23.Append(cell167);

                Row row24 = new Row() { RowIndex = (UInt32Value)12U, Spans = new ListValue<StringValue>() { InnerText = "1:3" }, DyDescent = 0.25D };

                Cell cell168 = new Cell() { CellReference = "A12", StyleIndex = (UInt32Value)3U, DataType = CellValues.SharedString };
                CellValue cellValue168 = new CellValue();
                cellValue168.Text = "36";

                cell168.Append(cellValue168);

                Cell cell169 = new Cell() { CellReference = "B12", StyleIndex = (UInt32Value)7U };
                CellValue cellValue169 = new CellValue();
                cellValue169.Text = "39236";

                cell169.Append(cellValue169);

                Cell cell170 = new Cell() { CellReference = "C12", StyleIndex = (UInt32Value)7U };
                CellValue cellValue170 = new CellValue();
                cellValue170.Text = "7847.1999999999989";

                cell170.Append(cellValue170);

                row24.Append(cell168);
                row24.Append(cell169);
                row24.Append(cell170);

                sheetData2.Append(row15);
                sheetData2.Append(row16);
                sheetData2.Append(row17);
                sheetData2.Append(row18);
                sheetData2.Append(row19);
                sheetData2.Append(row20);
                sheetData2.Append(row21);
                sheetData2.Append(row22);
                sheetData2.Append(row23);
                sheetData2.Append(row24);
                PageMargins pageMargins2 = new PageMargins() { Left = 0.7D, Right = 0.7D, Top = 0.75D, Bottom = 0.75D, Header = 0.3D, Footer = 0.3D };

                worksheet2.Append(sheetDimension2);
                worksheet2.Append(sheetViews2);
                worksheet2.Append(sheetFormatProperties2);
                worksheet2.Append(columns2);
                worksheet2.Append(sheetData2);
                worksheet2.Append(pageMargins2);

                worksheetPart2.Worksheet = worksheet2;
            }

            // Generates content of pivotTablePart1.
            private void GeneratePivotTablePart1Content(PivotTablePart pivotTablePart1)
            {
                PivotTableDefinition pivotTableDefinition1 = new PivotTableDefinition() { Name = "Сводная таблица1", CacheId = (UInt32Value)1U, ApplyNumberFormats = false, ApplyBorderFormats = false, ApplyFontFormats = false, ApplyPatternFormats = false, ApplyAlignmentFormats = false, ApplyWidthHeightFormats = true, DataCaption = "Значения", UpdatedVersion = 6, MinRefreshableVersion = 3, UseAutoFormatting = true, ItemPrintTitles = true, CreatedVersion = 6, Indent = (UInt32Value)0U, Outline = true, OutlineData = true, MultipleFieldFilters = false };
                Location location1 = new Location() { Reference = "A3:C12", FirstHeaderRow = (UInt32Value)0U, FirstDataRow = (UInt32Value)1U, FirstDataColumn = (UInt32Value)1U };

                PivotFields pivotFields1 = new PivotFields() { Count = (UInt32Value)10U };

                PivotField pivotField1 = new PivotField() { Axis = PivotTableAxisValues.AxisRow, ShowAll = false };

                Items items1 = new Items() { Count = (UInt32Value)4U };
                Item item1 = new Item() { Index = (UInt32Value)2U };
                Item item2 = new Item() { Index = (UInt32Value)0U };
                Item item3 = new Item() { Index = (UInt32Value)1U };
                Item item4 = new Item() { ItemType = ItemValues.Default };

                items1.Append(item1);
                items1.Append(item2);
                items1.Append(item3);
                items1.Append(item4);

                pivotField1.Append(items1);

                PivotField pivotField2 = new PivotField() { Axis = PivotTableAxisValues.AxisRow, ShowAll = false };

                Items items2 = new Items() { Count = (UInt32Value)4U };
                Item item5 = new Item() { Index = (UInt32Value)2U };
                Item item6 = new Item() { Index = (UInt32Value)0U };
                Item item7 = new Item() { Index = (UInt32Value)1U };
                Item item8 = new Item() { ItemType = ItemValues.Default };

                items2.Append(item5);
                items2.Append(item6);
                items2.Append(item7);
                items2.Append(item8);

                pivotField2.Append(items2);
                PivotField pivotField3 = new PivotField() { ShowAll = false };
                PivotField pivotField4 = new PivotField() { DataField = true, ShowAll = false };
                PivotField pivotField5 = new PivotField() { DataField = true, ShowAll = false };
                PivotField pivotField6 = new PivotField() { ShowAll = false };

                PivotField pivotField7 = new PivotField() { Axis = PivotTableAxisValues.AxisRow, ShowAll = false };

                Items items3 = new Items() { Count = (UInt32Value)2U };
                Item item9 = new Item() { Index = (UInt32Value)0U };
                Item item10 = new Item() { ItemType = ItemValues.Default };

                items3.Append(item9);
                items3.Append(item10);

                pivotField7.Append(items3);
                PivotField pivotField8 = new PivotField() { ShowAll = false };
                PivotField pivotField9 = new PivotField() { ShowAll = false };

                PivotField pivotField10 = new PivotField() { Axis = PivotTableAxisValues.AxisRow, ShowAll = false };

                Items items4 = new Items() { Count = (UInt32Value)2U };
                Item item11 = new Item() { Index = (UInt32Value)0U };
                Item item12 = new Item() { ItemType = ItemValues.Default };

                items4.Append(item11);
                items4.Append(item12);

                pivotField10.Append(items4);

                pivotFields1.Append(pivotField1);
                pivotFields1.Append(pivotField2);
                pivotFields1.Append(pivotField3);
                pivotFields1.Append(pivotField4);
                pivotFields1.Append(pivotField5);
                pivotFields1.Append(pivotField6);
                pivotFields1.Append(pivotField7);
                pivotFields1.Append(pivotField8);
                pivotFields1.Append(pivotField9);
                pivotFields1.Append(pivotField10);

                RowFields rowFields1 = new RowFields() { Count = (UInt32Value)4U };
                Field field1 = new Field() { Index = 9 };
                Field field2 = new Field() { Index = 6 };
                Field field3 = new Field() { Index = 1 };
                Field field4 = new Field() { Index = 0 };

                rowFields1.Append(field1);
                rowFields1.Append(field2);
                rowFields1.Append(field3);
                rowFields1.Append(field4);

                RowItems rowItems1 = new RowItems() { Count = (UInt32Value)9U };

                RowItem rowItem1 = new RowItem();
                MemberPropertyIndex memberPropertyIndex1 = new MemberPropertyIndex();

                rowItem1.Append(memberPropertyIndex1);

                RowItem rowItem2 = new RowItem() { RepeatedItemCount = (UInt32Value)1U };
                MemberPropertyIndex memberPropertyIndex2 = new MemberPropertyIndex();

                rowItem2.Append(memberPropertyIndex2);

                RowItem rowItem3 = new RowItem() { RepeatedItemCount = (UInt32Value)2U };
                MemberPropertyIndex memberPropertyIndex3 = new MemberPropertyIndex();

                rowItem3.Append(memberPropertyIndex3);

                RowItem rowItem4 = new RowItem() { RepeatedItemCount = (UInt32Value)3U };
                MemberPropertyIndex memberPropertyIndex4 = new MemberPropertyIndex();

                rowItem4.Append(memberPropertyIndex4);

                RowItem rowItem5 = new RowItem() { RepeatedItemCount = (UInt32Value)2U };
                MemberPropertyIndex memberPropertyIndex5 = new MemberPropertyIndex() { Val = 1 };

                rowItem5.Append(memberPropertyIndex5);

                RowItem rowItem6 = new RowItem() { RepeatedItemCount = (UInt32Value)3U };
                MemberPropertyIndex memberPropertyIndex6 = new MemberPropertyIndex() { Val = 1 };

                rowItem6.Append(memberPropertyIndex6);

                RowItem rowItem7 = new RowItem() { RepeatedItemCount = (UInt32Value)2U };
                MemberPropertyIndex memberPropertyIndex7 = new MemberPropertyIndex() { Val = 2 };

                rowItem7.Append(memberPropertyIndex7);

                RowItem rowItem8 = new RowItem() { RepeatedItemCount = (UInt32Value)3U };
                MemberPropertyIndex memberPropertyIndex8 = new MemberPropertyIndex() { Val = 2 };

                rowItem8.Append(memberPropertyIndex8);

                RowItem rowItem9 = new RowItem() { ItemType = ItemValues.Grand };
                MemberPropertyIndex memberPropertyIndex9 = new MemberPropertyIndex();

                rowItem9.Append(memberPropertyIndex9);

                rowItems1.Append(rowItem1);
                rowItems1.Append(rowItem2);
                rowItems1.Append(rowItem3);
                rowItems1.Append(rowItem4);
                rowItems1.Append(rowItem5);
                rowItems1.Append(rowItem6);
                rowItems1.Append(rowItem7);
                rowItems1.Append(rowItem8);
                rowItems1.Append(rowItem9);

                ColumnFields columnFields1 = new ColumnFields() { Count = (UInt32Value)1U };
                Field field5 = new Field() { Index = -2 };

                columnFields1.Append(field5);

                ColumnItems columnItems1 = new ColumnItems() { Count = (UInt32Value)2U };

                RowItem rowItem10 = new RowItem();
                MemberPropertyIndex memberPropertyIndex10 = new MemberPropertyIndex();

                rowItem10.Append(memberPropertyIndex10);

                RowItem rowItem11 = new RowItem() { Index = (UInt32Value)1U };
                MemberPropertyIndex memberPropertyIndex11 = new MemberPropertyIndex() { Val = 1 };

                rowItem11.Append(memberPropertyIndex11);

                columnItems1.Append(rowItem10);
                columnItems1.Append(rowItem11);

                DataFields dataFields1 = new DataFields() { Count = (UInt32Value)2U };
                DataField dataField1 = new DataField() { Name = "Сумма по полю СП", Field = (UInt32Value)3U, BaseField = 0, BaseItem = (UInt32Value)0U };
                DataField dataField2 = new DataField() { Name = "Сумма по полю АВ", Field = (UInt32Value)4U, BaseField = 0, BaseItem = (UInt32Value)0U };

                dataFields1.Append(dataField1);
                dataFields1.Append(dataField2);
                PivotTableStyle pivotTableStyle1 = new PivotTableStyle() { Name = "PivotStyleLight16", ShowRowHeaders = true, ShowColumnHeaders = true, ShowRowStripes = false, ShowColumnStripes = false, ShowLastColumn = true };

                PivotTableDefinitionExtensionList pivotTableDefinitionExtensionList1 = new PivotTableDefinitionExtensionList();

                PivotTableDefinitionExtension pivotTableDefinitionExtension1 = new PivotTableDefinitionExtension() { Uri = "{962EF5D1-5CA2-4c93-8EF4-DBF5C05439D2}" };
                pivotTableDefinitionExtension1.AddNamespaceDeclaration("x14", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main");

                X14.PivotTableDefinition pivotTableDefinition2 = new X14.PivotTableDefinition() { HideValuesRow = true };
                pivotTableDefinition2.AddNamespaceDeclaration("xm", "http://schemas.microsoft.com/office/excel/2006/main");

                pivotTableDefinitionExtension1.Append(pivotTableDefinition2);

                pivotTableDefinitionExtensionList1.Append(pivotTableDefinitionExtension1);

                pivotTableDefinition1.Append(location1);
                pivotTableDefinition1.Append(pivotFields1);
                pivotTableDefinition1.Append(rowFields1);
                pivotTableDefinition1.Append(rowItems1);
                pivotTableDefinition1.Append(columnFields1);
                pivotTableDefinition1.Append(columnItems1);
                pivotTableDefinition1.Append(dataFields1);
                pivotTableDefinition1.Append(pivotTableStyle1);
                pivotTableDefinition1.Append(pivotTableDefinitionExtensionList1);

                pivotTablePart1.PivotTableDefinition = pivotTableDefinition1;
            }

            // Generates content of sharedStringTablePart1.
            private void GenerateSharedStringTablePart1Content(SharedStringTablePart sharedStringTablePart1)
            {
                SharedStringTable sharedStringTable1 = new SharedStringTable() { Count = (UInt32Value)113U, UniqueCount = (UInt32Value)39U };

                SharedStringItem sharedStringItem1 = new SharedStringItem();
                Text text1 = new Text();
                text1.Text = "2844613435";

                sharedStringItem1.Append(text1);

                SharedStringItem sharedStringItem2 = new SharedStringItem();
                Text text2 = new Text();
                text2.Text = "Марусинець З.С.";

                sharedStringItem2.Append(text2);

                SharedStringItem sharedStringItem3 = new SharedStringItem();
                Text text3 = new Text();
                text3.Text = "231";

                sharedStringItem3.Append(text3);

                SharedStringItem sharedStringItem4 = new SharedStringItem();
                Text text4 = new Text();
                text4.Text = "0701";

                sharedStringItem4.Append(text4);

                SharedStringItem sharedStringItem5 = new SharedStringItem();
                Text text5 = new Text();
                text5.Text = "22 - фізичні особи - суб’єкти підприємницької діяльності (Категорія 1)";

                sharedStringItem5.Append(text5);

                SharedStringItem sharedStringItem6 = new SharedStringItem();
                Text text6 = new Text();
                text6.Text = "Агентський договір";

                sharedStringItem6.Append(text6);

                SharedStringItem sharedStringItem7 = new SharedStringItem();
                Text text7 = new Text();
                text7.Text = "501898";

                sharedStringItem7.Append(text7);

                SharedStringItem sharedStringItem8 = new SharedStringItem();
                Text text8 = new Text();
                text8.Text = "232";

                sharedStringItem8.Append(text8);

                SharedStringItem sharedStringItem9 = new SharedStringItem();
                Text text9 = new Text();
                text9.Text = "0703";

                sharedStringItem9.Append(text9);

                SharedStringItem sharedStringItem10 = new SharedStringItem();
                Text text10 = new Text();
                text10.Text = "0704";

                sharedStringItem10.Append(text10);

                SharedStringItem sharedStringItem11 = new SharedStringItem();
                Text text11 = new Text();
                text11.Text = "0706";

                sharedStringItem11.Append(text11);

                SharedStringItem sharedStringItem12 = new SharedStringItem();
                Text text12 = new Text();
                text12.Text = "0708";

                sharedStringItem12.Append(text12);

                SharedStringItem sharedStringItem13 = new SharedStringItem();
                Text text13 = new Text();
                text13.Text = "0711";

                sharedStringItem13.Append(text13);

                SharedStringItem sharedStringItem14 = new SharedStringItem();
                Text text14 = new Text();
                text14.Text = "3425509799";

                sharedStringItem14.Append(text14);

                SharedStringItem sharedStringItem15 = new SharedStringItem();
                Text text15 = new Text();
                text15.Text = "Мейсарош Е.Т.";

                sharedStringItem15.Append(text15);

                SharedStringItem sharedStringItem16 = new SharedStringItem();
                Text text16 = new Text();
                text16.Text = "0712";

                sharedStringItem16.Append(text16);

                SharedStringItem sharedStringItem17 = new SharedStringItem();
                Text text17 = new Text();
                text17.Text = "501897";

                sharedStringItem17.Append(text17);

                SharedStringItem sharedStringItem18 = new SharedStringItem();
                Text text18 = new Text();
                text18.Text = "2209911436";

                sharedStringItem18.Append(text18);

                SharedStringItem sharedStringItem19 = new SharedStringItem();
                Text text19 = new Text();
                text19.Text = "Гобан Ю.Ю.";

                sharedStringItem19.Append(text19);

                SharedStringItem sharedStringItem20 = new SharedStringItem();
                Text text20 = new Text();
                text20.Text = "0715";

                sharedStringItem20.Append(text20);

                SharedStringItem sharedStringItem21 = new SharedStringItem();
                Text text21 = new Text();
                text21.Text = "501896";

                sharedStringItem21.Append(text21);

                SharedStringItem sharedStringItem22 = new SharedStringItem();
                Text text22 = new Text();
                text22.Text = "0716";

                sharedStringItem22.Append(text22);

                SharedStringItem sharedStringItem23 = new SharedStringItem();
                Text text23 = new Text();
                text23.Text = "0718";

                sharedStringItem23.Append(text23);

                SharedStringItem sharedStringItem24 = new SharedStringItem();
                Text text24 = new Text();
                text24.Text = "0790";

                sharedStringItem24.Append(text24);

                SharedStringItem sharedStringItem25 = new SharedStringItem();
                Text text25 = new Text();
                text25.Text = "ІПН";

                sharedStringItem25.Append(text25);

                SharedStringItem sharedStringItem26 = new SharedStringItem();
                Text text26 = new Text();
                text26.Text = "Агент";

                sharedStringItem26.Append(text26);

                SharedStringItem sharedStringItem27 = new SharedStringItem();
                Text text27 = new Text();
                text27.Text = "Код програми";

                sharedStringItem27.Append(text27);

                SharedStringItem sharedStringItem28 = new SharedStringItem();
                Text text28 = new Text();
                text28.Text = "СП";

                sharedStringItem28.Append(text28);

                SharedStringItem sharedStringItem29 = new SharedStringItem();
                Text text29 = new Text();
                text29.Text = "АВ";

                sharedStringItem29.Append(text29);

                SharedStringItem sharedStringItem30 = new SharedStringItem();
                Text text30 = new Text();
                text30.Text = "Код відділення";

                sharedStringItem30.Append(text30);

                SharedStringItem sharedStringItem31 = new SharedStringItem();
                Text text31 = new Text();
                text31.Text = "Канал";

                sharedStringItem31.Append(text31);

                SharedStringItem sharedStringItem32 = new SharedStringItem();
                Text text32 = new Text();
                text32.Text = "Договір";

                sharedStringItem32.Append(text32);

                SharedStringItem sharedStringItem33 = new SharedStringItem();
                Text text33 = new Text();
                text33.Text = "ID акту";

                sharedStringItem33.Append(text33);

                SharedStringItem sharedStringItem34 = new SharedStringItem();
                Text text34 = new Text();
                text34.Text = "Дирекція";

                sharedStringItem34.Append(text34);

                SharedStringItem sharedStringItem35 = new SharedStringItem();
                Text text35 = new Text();
                text35.Text = "Названия строк";

                sharedStringItem35.Append(text35);

                SharedStringItem sharedStringItem36 = new SharedStringItem();
                Text text36 = new Text();
                text36.Text = "07";

                sharedStringItem36.Append(text36);

                SharedStringItem sharedStringItem37 = new SharedStringItem();
                Text text37 = new Text();
                text37.Text = "Общий итог";

                sharedStringItem37.Append(text37);

                SharedStringItem sharedStringItem38 = new SharedStringItem();
                Text text38 = new Text();
                text38.Text = "Сумма по полю СП";

                sharedStringItem38.Append(text38);

                SharedStringItem sharedStringItem39 = new SharedStringItem();
                Text text39 = new Text();
                text39.Text = "Сумма по полю АВ";

                sharedStringItem39.Append(text39);

                sharedStringTable1.Append(sharedStringItem1);
                sharedStringTable1.Append(sharedStringItem2);
                sharedStringTable1.Append(sharedStringItem3);
                sharedStringTable1.Append(sharedStringItem4);
                sharedStringTable1.Append(sharedStringItem5);
                sharedStringTable1.Append(sharedStringItem6);
                sharedStringTable1.Append(sharedStringItem7);
                sharedStringTable1.Append(sharedStringItem8);
                sharedStringTable1.Append(sharedStringItem9);
                sharedStringTable1.Append(sharedStringItem10);
                sharedStringTable1.Append(sharedStringItem11);
                sharedStringTable1.Append(sharedStringItem12);
                sharedStringTable1.Append(sharedStringItem13);
                sharedStringTable1.Append(sharedStringItem14);
                sharedStringTable1.Append(sharedStringItem15);
                sharedStringTable1.Append(sharedStringItem16);
                sharedStringTable1.Append(sharedStringItem17);
                sharedStringTable1.Append(sharedStringItem18);
                sharedStringTable1.Append(sharedStringItem19);
                sharedStringTable1.Append(sharedStringItem20);
                sharedStringTable1.Append(sharedStringItem21);
                sharedStringTable1.Append(sharedStringItem22);
                sharedStringTable1.Append(sharedStringItem23);
                sharedStringTable1.Append(sharedStringItem24);
                sharedStringTable1.Append(sharedStringItem25);
                sharedStringTable1.Append(sharedStringItem26);
                sharedStringTable1.Append(sharedStringItem27);
                sharedStringTable1.Append(sharedStringItem28);
                sharedStringTable1.Append(sharedStringItem29);
                sharedStringTable1.Append(sharedStringItem30);
                sharedStringTable1.Append(sharedStringItem31);
                sharedStringTable1.Append(sharedStringItem32);
                sharedStringTable1.Append(sharedStringItem33);
                sharedStringTable1.Append(sharedStringItem34);
                sharedStringTable1.Append(sharedStringItem35);
                sharedStringTable1.Append(sharedStringItem36);
                sharedStringTable1.Append(sharedStringItem37);
                sharedStringTable1.Append(sharedStringItem38);
                sharedStringTable1.Append(sharedStringItem39);

                sharedStringTablePart1.SharedStringTable = sharedStringTable1;
            }

            // Generates content of workbookStylesPart1.
            private void GenerateWorkbookStylesPart1Content(WorkbookStylesPart workbookStylesPart1)
            {
                Stylesheet stylesheet1 = new Stylesheet() { MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "x14ac x16r2" } };
                stylesheet1.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
                stylesheet1.AddNamespaceDeclaration("x14ac", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac");
                stylesheet1.AddNamespaceDeclaration("x16r2", "http://schemas.microsoft.com/office/spreadsheetml/2015/02/main");

                Fonts fonts1 = new Fonts() { Count = (UInt32Value)18U, KnownFonts = true };

                Font font1 = new Font();
                FontSize fontSize1 = new FontSize() { Val = 11D };
                Color color1 = new Color() { Theme = (UInt32Value)1U };
                FontName fontName1 = new FontName() { Val = "Calibri" };
                FontFamilyNumbering fontFamilyNumbering1 = new FontFamilyNumbering() { Val = 2 };
                FontCharSet fontCharSet1 = new FontCharSet() { Val = 204 };
                FontScheme fontScheme1 = new FontScheme() { Val = FontSchemeValues.Minor };

                font1.Append(fontSize1);
                font1.Append(color1);
                font1.Append(fontName1);
                font1.Append(fontFamilyNumbering1);
                font1.Append(fontCharSet1);
                font1.Append(fontScheme1);

                Font font2 = new Font();
                FontSize fontSize2 = new FontSize() { Val = 11D };
                Color color2 = new Color() { Theme = (UInt32Value)1U };
                FontName fontName2 = new FontName() { Val = "Calibri" };
                FontFamilyNumbering fontFamilyNumbering2 = new FontFamilyNumbering() { Val = 2 };
                FontCharSet fontCharSet2 = new FontCharSet() { Val = 204 };
                FontScheme fontScheme2 = new FontScheme() { Val = FontSchemeValues.Minor };

                font2.Append(fontSize2);
                font2.Append(color2);
                font2.Append(fontName2);
                font2.Append(fontFamilyNumbering2);
                font2.Append(fontCharSet2);
                font2.Append(fontScheme2);

                Font font3 = new Font();
                FontSize fontSize3 = new FontSize() { Val = 18D };
                Color color3 = new Color() { Theme = (UInt32Value)3U };
                FontName fontName3 = new FontName() { Val = "Calibri Light" };
                FontFamilyNumbering fontFamilyNumbering3 = new FontFamilyNumbering() { Val = 2 };
                FontCharSet fontCharSet3 = new FontCharSet() { Val = 204 };
                FontScheme fontScheme3 = new FontScheme() { Val = FontSchemeValues.Major };

                font3.Append(fontSize3);
                font3.Append(color3);
                font3.Append(fontName3);
                font3.Append(fontFamilyNumbering3);
                font3.Append(fontCharSet3);
                font3.Append(fontScheme3);

                Font font4 = new Font();
                Bold bold1 = new Bold();
                FontSize fontSize4 = new FontSize() { Val = 15D };
                Color color4 = new Color() { Theme = (UInt32Value)3U };
                FontName fontName4 = new FontName() { Val = "Calibri" };
                FontFamilyNumbering fontFamilyNumbering4 = new FontFamilyNumbering() { Val = 2 };
                FontCharSet fontCharSet4 = new FontCharSet() { Val = 204 };
                FontScheme fontScheme4 = new FontScheme() { Val = FontSchemeValues.Minor };

                font4.Append(bold1);
                font4.Append(fontSize4);
                font4.Append(color4);
                font4.Append(fontName4);
                font4.Append(fontFamilyNumbering4);
                font4.Append(fontCharSet4);
                font4.Append(fontScheme4);

                Font font5 = new Font();
                Bold bold2 = new Bold();
                FontSize fontSize5 = new FontSize() { Val = 13D };
                Color color5 = new Color() { Theme = (UInt32Value)3U };
                FontName fontName5 = new FontName() { Val = "Calibri" };
                FontFamilyNumbering fontFamilyNumbering5 = new FontFamilyNumbering() { Val = 2 };
                FontCharSet fontCharSet5 = new FontCharSet() { Val = 204 };
                FontScheme fontScheme5 = new FontScheme() { Val = FontSchemeValues.Minor };

                font5.Append(bold2);
                font5.Append(fontSize5);
                font5.Append(color5);
                font5.Append(fontName5);
                font5.Append(fontFamilyNumbering5);
                font5.Append(fontCharSet5);
                font5.Append(fontScheme5);

                Font font6 = new Font();
                Bold bold3 = new Bold();
                FontSize fontSize6 = new FontSize() { Val = 11D };
                Color color6 = new Color() { Theme = (UInt32Value)3U };
                FontName fontName6 = new FontName() { Val = "Calibri" };
                FontFamilyNumbering fontFamilyNumbering6 = new FontFamilyNumbering() { Val = 2 };
                FontCharSet fontCharSet6 = new FontCharSet() { Val = 204 };
                FontScheme fontScheme6 = new FontScheme() { Val = FontSchemeValues.Minor };

                font6.Append(bold3);
                font6.Append(fontSize6);
                font6.Append(color6);
                font6.Append(fontName6);
                font6.Append(fontFamilyNumbering6);
                font6.Append(fontCharSet6);
                font6.Append(fontScheme6);

                Font font7 = new Font();
                FontSize fontSize7 = new FontSize() { Val = 11D };
                Color color7 = new Color() { Rgb = "FF006100" };
                FontName fontName7 = new FontName() { Val = "Calibri" };
                FontFamilyNumbering fontFamilyNumbering7 = new FontFamilyNumbering() { Val = 2 };
                FontCharSet fontCharSet7 = new FontCharSet() { Val = 204 };
                FontScheme fontScheme7 = new FontScheme() { Val = FontSchemeValues.Minor };

                font7.Append(fontSize7);
                font7.Append(color7);
                font7.Append(fontName7);
                font7.Append(fontFamilyNumbering7);
                font7.Append(fontCharSet7);
                font7.Append(fontScheme7);

                Font font8 = new Font();
                FontSize fontSize8 = new FontSize() { Val = 11D };
                Color color8 = new Color() { Rgb = "FF9C0006" };
                FontName fontName8 = new FontName() { Val = "Calibri" };
                FontFamilyNumbering fontFamilyNumbering8 = new FontFamilyNumbering() { Val = 2 };
                FontCharSet fontCharSet8 = new FontCharSet() { Val = 204 };
                FontScheme fontScheme8 = new FontScheme() { Val = FontSchemeValues.Minor };

                font8.Append(fontSize8);
                font8.Append(color8);
                font8.Append(fontName8);
                font8.Append(fontFamilyNumbering8);
                font8.Append(fontCharSet8);
                font8.Append(fontScheme8);

                Font font9 = new Font();
                FontSize fontSize9 = new FontSize() { Val = 11D };
                Color color9 = new Color() { Rgb = "FF9C6500" };
                FontName fontName9 = new FontName() { Val = "Calibri" };
                FontFamilyNumbering fontFamilyNumbering9 = new FontFamilyNumbering() { Val = 2 };
                FontCharSet fontCharSet9 = new FontCharSet() { Val = 204 };
                FontScheme fontScheme9 = new FontScheme() { Val = FontSchemeValues.Minor };

                font9.Append(fontSize9);
                font9.Append(color9);
                font9.Append(fontName9);
                font9.Append(fontFamilyNumbering9);
                font9.Append(fontCharSet9);
                font9.Append(fontScheme9);

                Font font10 = new Font();
                FontSize fontSize10 = new FontSize() { Val = 11D };
                Color color10 = new Color() { Rgb = "FF3F3F76" };
                FontName fontName10 = new FontName() { Val = "Calibri" };
                FontFamilyNumbering fontFamilyNumbering10 = new FontFamilyNumbering() { Val = 2 };
                FontCharSet fontCharSet10 = new FontCharSet() { Val = 204 };
                FontScheme fontScheme10 = new FontScheme() { Val = FontSchemeValues.Minor };

                font10.Append(fontSize10);
                font10.Append(color10);
                font10.Append(fontName10);
                font10.Append(fontFamilyNumbering10);
                font10.Append(fontCharSet10);
                font10.Append(fontScheme10);

                Font font11 = new Font();
                Bold bold4 = new Bold();
                FontSize fontSize11 = new FontSize() { Val = 11D };
                Color color11 = new Color() { Rgb = "FF3F3F3F" };
                FontName fontName11 = new FontName() { Val = "Calibri" };
                FontFamilyNumbering fontFamilyNumbering11 = new FontFamilyNumbering() { Val = 2 };
                FontCharSet fontCharSet11 = new FontCharSet() { Val = 204 };
                FontScheme fontScheme11 = new FontScheme() { Val = FontSchemeValues.Minor };

                font11.Append(bold4);
                font11.Append(fontSize11);
                font11.Append(color11);
                font11.Append(fontName11);
                font11.Append(fontFamilyNumbering11);
                font11.Append(fontCharSet11);
                font11.Append(fontScheme11);

                Font font12 = new Font();
                Bold bold5 = new Bold();
                FontSize fontSize12 = new FontSize() { Val = 11D };
                Color color12 = new Color() { Rgb = "FFFA7D00" };
                FontName fontName12 = new FontName() { Val = "Calibri" };
                FontFamilyNumbering fontFamilyNumbering12 = new FontFamilyNumbering() { Val = 2 };
                FontCharSet fontCharSet12 = new FontCharSet() { Val = 204 };
                FontScheme fontScheme12 = new FontScheme() { Val = FontSchemeValues.Minor };

                font12.Append(bold5);
                font12.Append(fontSize12);
                font12.Append(color12);
                font12.Append(fontName12);
                font12.Append(fontFamilyNumbering12);
                font12.Append(fontCharSet12);
                font12.Append(fontScheme12);

                Font font13 = new Font();
                FontSize fontSize13 = new FontSize() { Val = 11D };
                Color color13 = new Color() { Rgb = "FFFA7D00" };
                FontName fontName13 = new FontName() { Val = "Calibri" };
                FontFamilyNumbering fontFamilyNumbering13 = new FontFamilyNumbering() { Val = 2 };
                FontCharSet fontCharSet13 = new FontCharSet() { Val = 204 };
                FontScheme fontScheme13 = new FontScheme() { Val = FontSchemeValues.Minor };

                font13.Append(fontSize13);
                font13.Append(color13);
                font13.Append(fontName13);
                font13.Append(fontFamilyNumbering13);
                font13.Append(fontCharSet13);
                font13.Append(fontScheme13);

                Font font14 = new Font();
                Bold bold6 = new Bold();
                FontSize fontSize14 = new FontSize() { Val = 11D };
                Color color14 = new Color() { Theme = (UInt32Value)0U };
                FontName fontName14 = new FontName() { Val = "Calibri" };
                FontFamilyNumbering fontFamilyNumbering14 = new FontFamilyNumbering() { Val = 2 };
                FontCharSet fontCharSet14 = new FontCharSet() { Val = 204 };
                FontScheme fontScheme14 = new FontScheme() { Val = FontSchemeValues.Minor };

                font14.Append(bold6);
                font14.Append(fontSize14);
                font14.Append(color14);
                font14.Append(fontName14);
                font14.Append(fontFamilyNumbering14);
                font14.Append(fontCharSet14);
                font14.Append(fontScheme14);

                Font font15 = new Font();
                FontSize fontSize15 = new FontSize() { Val = 11D };
                Color color15 = new Color() { Rgb = "FFFF0000" };
                FontName fontName15 = new FontName() { Val = "Calibri" };
                FontFamilyNumbering fontFamilyNumbering15 = new FontFamilyNumbering() { Val = 2 };
                FontCharSet fontCharSet15 = new FontCharSet() { Val = 204 };
                FontScheme fontScheme15 = new FontScheme() { Val = FontSchemeValues.Minor };

                font15.Append(fontSize15);
                font15.Append(color15);
                font15.Append(fontName15);
                font15.Append(fontFamilyNumbering15);
                font15.Append(fontCharSet15);
                font15.Append(fontScheme15);

                Font font16 = new Font();
                Italic italic1 = new Italic();
                FontSize fontSize16 = new FontSize() { Val = 11D };
                Color color16 = new Color() { Rgb = "FF7F7F7F" };
                FontName fontName16 = new FontName() { Val = "Calibri" };
                FontFamilyNumbering fontFamilyNumbering16 = new FontFamilyNumbering() { Val = 2 };
                FontCharSet fontCharSet16 = new FontCharSet() { Val = 204 };
                FontScheme fontScheme16 = new FontScheme() { Val = FontSchemeValues.Minor };

                font16.Append(italic1);
                font16.Append(fontSize16);
                font16.Append(color16);
                font16.Append(fontName16);
                font16.Append(fontFamilyNumbering16);
                font16.Append(fontCharSet16);
                font16.Append(fontScheme16);

                Font font17 = new Font();
                Bold bold7 = new Bold();
                FontSize fontSize17 = new FontSize() { Val = 11D };
                Color color17 = new Color() { Theme = (UInt32Value)1U };
                FontName fontName17 = new FontName() { Val = "Calibri" };
                FontFamilyNumbering fontFamilyNumbering17 = new FontFamilyNumbering() { Val = 2 };
                FontCharSet fontCharSet17 = new FontCharSet() { Val = 204 };
                FontScheme fontScheme17 = new FontScheme() { Val = FontSchemeValues.Minor };

                font17.Append(bold7);
                font17.Append(fontSize17);
                font17.Append(color17);
                font17.Append(fontName17);
                font17.Append(fontFamilyNumbering17);
                font17.Append(fontCharSet17);
                font17.Append(fontScheme17);

                Font font18 = new Font();
                FontSize fontSize18 = new FontSize() { Val = 11D };
                Color color18 = new Color() { Theme = (UInt32Value)0U };
                FontName fontName18 = new FontName() { Val = "Calibri" };
                FontFamilyNumbering fontFamilyNumbering18 = new FontFamilyNumbering() { Val = 2 };
                FontCharSet fontCharSet18 = new FontCharSet() { Val = 204 };
                FontScheme fontScheme18 = new FontScheme() { Val = FontSchemeValues.Minor };

                font18.Append(fontSize18);
                font18.Append(color18);
                font18.Append(fontName18);
                font18.Append(fontFamilyNumbering18);
                font18.Append(fontCharSet18);
                font18.Append(fontScheme18);

                fonts1.Append(font1);
                fonts1.Append(font2);
                fonts1.Append(font3);
                fonts1.Append(font4);
                fonts1.Append(font5);
                fonts1.Append(font6);
                fonts1.Append(font7);
                fonts1.Append(font8);
                fonts1.Append(font9);
                fonts1.Append(font10);
                fonts1.Append(font11);
                fonts1.Append(font12);
                fonts1.Append(font13);
                fonts1.Append(font14);
                fonts1.Append(font15);
                fonts1.Append(font16);
                fonts1.Append(font17);
                fonts1.Append(font18);

                Fills fills1 = new Fills() { Count = (UInt32Value)33U };

                Fill fill1 = new Fill();
                PatternFill patternFill1 = new PatternFill() { PatternType = PatternValues.None };

                fill1.Append(patternFill1);

                Fill fill2 = new Fill();
                PatternFill patternFill2 = new PatternFill() { PatternType = PatternValues.Gray125 };

                fill2.Append(patternFill2);

                Fill fill3 = new Fill();

                PatternFill patternFill3 = new PatternFill() { PatternType = PatternValues.Solid };
                ForegroundColor foregroundColor1 = new ForegroundColor() { Rgb = "FFC6EFCE" };

                patternFill3.Append(foregroundColor1);

                fill3.Append(patternFill3);

                Fill fill4 = new Fill();

                PatternFill patternFill4 = new PatternFill() { PatternType = PatternValues.Solid };
                ForegroundColor foregroundColor2 = new ForegroundColor() { Rgb = "FFFFC7CE" };

                patternFill4.Append(foregroundColor2);

                fill4.Append(patternFill4);

                Fill fill5 = new Fill();

                PatternFill patternFill5 = new PatternFill() { PatternType = PatternValues.Solid };
                ForegroundColor foregroundColor3 = new ForegroundColor() { Rgb = "FFFFEB9C" };

                patternFill5.Append(foregroundColor3);

                fill5.Append(patternFill5);

                Fill fill6 = new Fill();

                PatternFill patternFill6 = new PatternFill() { PatternType = PatternValues.Solid };
                ForegroundColor foregroundColor4 = new ForegroundColor() { Rgb = "FFFFCC99" };

                patternFill6.Append(foregroundColor4);

                fill6.Append(patternFill6);

                Fill fill7 = new Fill();

                PatternFill patternFill7 = new PatternFill() { PatternType = PatternValues.Solid };
                ForegroundColor foregroundColor5 = new ForegroundColor() { Rgb = "FFF2F2F2" };

                patternFill7.Append(foregroundColor5);

                fill7.Append(patternFill7);

                Fill fill8 = new Fill();

                PatternFill patternFill8 = new PatternFill() { PatternType = PatternValues.Solid };
                ForegroundColor foregroundColor6 = new ForegroundColor() { Rgb = "FFA5A5A5" };

                patternFill8.Append(foregroundColor6);

                fill8.Append(patternFill8);

                Fill fill9 = new Fill();

                PatternFill patternFill9 = new PatternFill() { PatternType = PatternValues.Solid };
                ForegroundColor foregroundColor7 = new ForegroundColor() { Rgb = "FFFFFFCC" };

                patternFill9.Append(foregroundColor7);

                fill9.Append(patternFill9);

                Fill fill10 = new Fill();

                PatternFill patternFill10 = new PatternFill() { PatternType = PatternValues.Solid };
                ForegroundColor foregroundColor8 = new ForegroundColor() { Theme = (UInt32Value)4U };

                patternFill10.Append(foregroundColor8);

                fill10.Append(patternFill10);

                Fill fill11 = new Fill();

                PatternFill patternFill11 = new PatternFill() { PatternType = PatternValues.Solid };
                ForegroundColor foregroundColor9 = new ForegroundColor() { Theme = (UInt32Value)4U, Tint = 0.79998168889431442D };
                BackgroundColor backgroundColor1 = new BackgroundColor() { Indexed = (UInt32Value)65U };

                patternFill11.Append(foregroundColor9);
                patternFill11.Append(backgroundColor1);

                fill11.Append(patternFill11);

                Fill fill12 = new Fill();

                PatternFill patternFill12 = new PatternFill() { PatternType = PatternValues.Solid };
                ForegroundColor foregroundColor10 = new ForegroundColor() { Theme = (UInt32Value)4U, Tint = 0.59999389629810485D };
                BackgroundColor backgroundColor2 = new BackgroundColor() { Indexed = (UInt32Value)65U };

                patternFill12.Append(foregroundColor10);
                patternFill12.Append(backgroundColor2);

                fill12.Append(patternFill12);

                Fill fill13 = new Fill();

                PatternFill patternFill13 = new PatternFill() { PatternType = PatternValues.Solid };
                ForegroundColor foregroundColor11 = new ForegroundColor() { Theme = (UInt32Value)4U, Tint = 0.39997558519241921D };
                BackgroundColor backgroundColor3 = new BackgroundColor() { Indexed = (UInt32Value)65U };

                patternFill13.Append(foregroundColor11);
                patternFill13.Append(backgroundColor3);

                fill13.Append(patternFill13);

                Fill fill14 = new Fill();

                PatternFill patternFill14 = new PatternFill() { PatternType = PatternValues.Solid };
                ForegroundColor foregroundColor12 = new ForegroundColor() { Theme = (UInt32Value)5U };

                patternFill14.Append(foregroundColor12);

                fill14.Append(patternFill14);

                Fill fill15 = new Fill();

                PatternFill patternFill15 = new PatternFill() { PatternType = PatternValues.Solid };
                ForegroundColor foregroundColor13 = new ForegroundColor() { Theme = (UInt32Value)5U, Tint = 0.79998168889431442D };
                BackgroundColor backgroundColor4 = new BackgroundColor() { Indexed = (UInt32Value)65U };

                patternFill15.Append(foregroundColor13);
                patternFill15.Append(backgroundColor4);

                fill15.Append(patternFill15);

                Fill fill16 = new Fill();

                PatternFill patternFill16 = new PatternFill() { PatternType = PatternValues.Solid };
                ForegroundColor foregroundColor14 = new ForegroundColor() { Theme = (UInt32Value)5U, Tint = 0.59999389629810485D };
                BackgroundColor backgroundColor5 = new BackgroundColor() { Indexed = (UInt32Value)65U };

                patternFill16.Append(foregroundColor14);
                patternFill16.Append(backgroundColor5);

                fill16.Append(patternFill16);

                Fill fill17 = new Fill();

                PatternFill patternFill17 = new PatternFill() { PatternType = PatternValues.Solid };
                ForegroundColor foregroundColor15 = new ForegroundColor() { Theme = (UInt32Value)5U, Tint = 0.39997558519241921D };
                BackgroundColor backgroundColor6 = new BackgroundColor() { Indexed = (UInt32Value)65U };

                patternFill17.Append(foregroundColor15);
                patternFill17.Append(backgroundColor6);

                fill17.Append(patternFill17);

                Fill fill18 = new Fill();

                PatternFill patternFill18 = new PatternFill() { PatternType = PatternValues.Solid };
                ForegroundColor foregroundColor16 = new ForegroundColor() { Theme = (UInt32Value)6U };

                patternFill18.Append(foregroundColor16);

                fill18.Append(patternFill18);

                Fill fill19 = new Fill();

                PatternFill patternFill19 = new PatternFill() { PatternType = PatternValues.Solid };
                ForegroundColor foregroundColor17 = new ForegroundColor() { Theme = (UInt32Value)6U, Tint = 0.79998168889431442D };
                BackgroundColor backgroundColor7 = new BackgroundColor() { Indexed = (UInt32Value)65U };

                patternFill19.Append(foregroundColor17);
                patternFill19.Append(backgroundColor7);

                fill19.Append(patternFill19);

                Fill fill20 = new Fill();

                PatternFill patternFill20 = new PatternFill() { PatternType = PatternValues.Solid };
                ForegroundColor foregroundColor18 = new ForegroundColor() { Theme = (UInt32Value)6U, Tint = 0.59999389629810485D };
                BackgroundColor backgroundColor8 = new BackgroundColor() { Indexed = (UInt32Value)65U };

                patternFill20.Append(foregroundColor18);
                patternFill20.Append(backgroundColor8);

                fill20.Append(patternFill20);

                Fill fill21 = new Fill();

                PatternFill patternFill21 = new PatternFill() { PatternType = PatternValues.Solid };
                ForegroundColor foregroundColor19 = new ForegroundColor() { Theme = (UInt32Value)6U, Tint = 0.39997558519241921D };
                BackgroundColor backgroundColor9 = new BackgroundColor() { Indexed = (UInt32Value)65U };

                patternFill21.Append(foregroundColor19);
                patternFill21.Append(backgroundColor9);

                fill21.Append(patternFill21);

                Fill fill22 = new Fill();

                PatternFill patternFill22 = new PatternFill() { PatternType = PatternValues.Solid };
                ForegroundColor foregroundColor20 = new ForegroundColor() { Theme = (UInt32Value)7U };

                patternFill22.Append(foregroundColor20);

                fill22.Append(patternFill22);

                Fill fill23 = new Fill();

                PatternFill patternFill23 = new PatternFill() { PatternType = PatternValues.Solid };
                ForegroundColor foregroundColor21 = new ForegroundColor() { Theme = (UInt32Value)7U, Tint = 0.79998168889431442D };
                BackgroundColor backgroundColor10 = new BackgroundColor() { Indexed = (UInt32Value)65U };

                patternFill23.Append(foregroundColor21);
                patternFill23.Append(backgroundColor10);

                fill23.Append(patternFill23);

                Fill fill24 = new Fill();

                PatternFill patternFill24 = new PatternFill() { PatternType = PatternValues.Solid };
                ForegroundColor foregroundColor22 = new ForegroundColor() { Theme = (UInt32Value)7U, Tint = 0.59999389629810485D };
                BackgroundColor backgroundColor11 = new BackgroundColor() { Indexed = (UInt32Value)65U };

                patternFill24.Append(foregroundColor22);
                patternFill24.Append(backgroundColor11);

                fill24.Append(patternFill24);

                Fill fill25 = new Fill();

                PatternFill patternFill25 = new PatternFill() { PatternType = PatternValues.Solid };
                ForegroundColor foregroundColor23 = new ForegroundColor() { Theme = (UInt32Value)7U, Tint = 0.39997558519241921D };
                BackgroundColor backgroundColor12 = new BackgroundColor() { Indexed = (UInt32Value)65U };

                patternFill25.Append(foregroundColor23);
                patternFill25.Append(backgroundColor12);

                fill25.Append(patternFill25);

                Fill fill26 = new Fill();

                PatternFill patternFill26 = new PatternFill() { PatternType = PatternValues.Solid };
                ForegroundColor foregroundColor24 = new ForegroundColor() { Theme = (UInt32Value)8U };

                patternFill26.Append(foregroundColor24);

                fill26.Append(patternFill26);

                Fill fill27 = new Fill();

                PatternFill patternFill27 = new PatternFill() { PatternType = PatternValues.Solid };
                ForegroundColor foregroundColor25 = new ForegroundColor() { Theme = (UInt32Value)8U, Tint = 0.79998168889431442D };
                BackgroundColor backgroundColor13 = new BackgroundColor() { Indexed = (UInt32Value)65U };

                patternFill27.Append(foregroundColor25);
                patternFill27.Append(backgroundColor13);

                fill27.Append(patternFill27);

                Fill fill28 = new Fill();

                PatternFill patternFill28 = new PatternFill() { PatternType = PatternValues.Solid };
                ForegroundColor foregroundColor26 = new ForegroundColor() { Theme = (UInt32Value)8U, Tint = 0.59999389629810485D };
                BackgroundColor backgroundColor14 = new BackgroundColor() { Indexed = (UInt32Value)65U };

                patternFill28.Append(foregroundColor26);
                patternFill28.Append(backgroundColor14);

                fill28.Append(patternFill28);

                Fill fill29 = new Fill();

                PatternFill patternFill29 = new PatternFill() { PatternType = PatternValues.Solid };
                ForegroundColor foregroundColor27 = new ForegroundColor() { Theme = (UInt32Value)8U, Tint = 0.39997558519241921D };
                BackgroundColor backgroundColor15 = new BackgroundColor() { Indexed = (UInt32Value)65U };

                patternFill29.Append(foregroundColor27);
                patternFill29.Append(backgroundColor15);

                fill29.Append(patternFill29);

                Fill fill30 = new Fill();

                PatternFill patternFill30 = new PatternFill() { PatternType = PatternValues.Solid };
                ForegroundColor foregroundColor28 = new ForegroundColor() { Theme = (UInt32Value)9U };

                patternFill30.Append(foregroundColor28);

                fill30.Append(patternFill30);

                Fill fill31 = new Fill();

                PatternFill patternFill31 = new PatternFill() { PatternType = PatternValues.Solid };
                ForegroundColor foregroundColor29 = new ForegroundColor() { Theme = (UInt32Value)9U, Tint = 0.79998168889431442D };
                BackgroundColor backgroundColor16 = new BackgroundColor() { Indexed = (UInt32Value)65U };

                patternFill31.Append(foregroundColor29);
                patternFill31.Append(backgroundColor16);

                fill31.Append(patternFill31);

                Fill fill32 = new Fill();

                PatternFill patternFill32 = new PatternFill() { PatternType = PatternValues.Solid };
                ForegroundColor foregroundColor30 = new ForegroundColor() { Theme = (UInt32Value)9U, Tint = 0.59999389629810485D };
                BackgroundColor backgroundColor17 = new BackgroundColor() { Indexed = (UInt32Value)65U };

                patternFill32.Append(foregroundColor30);
                patternFill32.Append(backgroundColor17);

                fill32.Append(patternFill32);

                Fill fill33 = new Fill();

                PatternFill patternFill33 = new PatternFill() { PatternType = PatternValues.Solid };
                ForegroundColor foregroundColor31 = new ForegroundColor() { Theme = (UInt32Value)9U, Tint = 0.39997558519241921D };
                BackgroundColor backgroundColor18 = new BackgroundColor() { Indexed = (UInt32Value)65U };

                patternFill33.Append(foregroundColor31);
                patternFill33.Append(backgroundColor18);

                fill33.Append(patternFill33);

                fills1.Append(fill1);
                fills1.Append(fill2);
                fills1.Append(fill3);
                fills1.Append(fill4);
                fills1.Append(fill5);
                fills1.Append(fill6);
                fills1.Append(fill7);
                fills1.Append(fill8);
                fills1.Append(fill9);
                fills1.Append(fill10);
                fills1.Append(fill11);
                fills1.Append(fill12);
                fills1.Append(fill13);
                fills1.Append(fill14);
                fills1.Append(fill15);
                fills1.Append(fill16);
                fills1.Append(fill17);
                fills1.Append(fill18);
                fills1.Append(fill19);
                fills1.Append(fill20);
                fills1.Append(fill21);
                fills1.Append(fill22);
                fills1.Append(fill23);
                fills1.Append(fill24);
                fills1.Append(fill25);
                fills1.Append(fill26);
                fills1.Append(fill27);
                fills1.Append(fill28);
                fills1.Append(fill29);
                fills1.Append(fill30);
                fills1.Append(fill31);
                fills1.Append(fill32);
                fills1.Append(fill33);

                Borders borders1 = new Borders() { Count = (UInt32Value)10U };

                Border border1 = new Border();
                LeftBorder leftBorder1 = new LeftBorder();
                RightBorder rightBorder1 = new RightBorder();
                TopBorder topBorder1 = new TopBorder();
                BottomBorder bottomBorder1 = new BottomBorder();
                DiagonalBorder diagonalBorder1 = new DiagonalBorder();

                border1.Append(leftBorder1);
                border1.Append(rightBorder1);
                border1.Append(topBorder1);
                border1.Append(bottomBorder1);
                border1.Append(diagonalBorder1);

                Border border2 = new Border();
                LeftBorder leftBorder2 = new LeftBorder();
                RightBorder rightBorder2 = new RightBorder();
                TopBorder topBorder2 = new TopBorder();

                BottomBorder bottomBorder2 = new BottomBorder() { Style = BorderStyleValues.Thick };
                Color color19 = new Color() { Theme = (UInt32Value)4U };

                bottomBorder2.Append(color19);
                DiagonalBorder diagonalBorder2 = new DiagonalBorder();

                border2.Append(leftBorder2);
                border2.Append(rightBorder2);
                border2.Append(topBorder2);
                border2.Append(bottomBorder2);
                border2.Append(diagonalBorder2);

                Border border3 = new Border();
                LeftBorder leftBorder3 = new LeftBorder();
                RightBorder rightBorder3 = new RightBorder();
                TopBorder topBorder3 = new TopBorder();

                BottomBorder bottomBorder3 = new BottomBorder() { Style = BorderStyleValues.Thick };
                Color color20 = new Color() { Theme = (UInt32Value)4U, Tint = 0.499984740745262D };

                bottomBorder3.Append(color20);
                DiagonalBorder diagonalBorder3 = new DiagonalBorder();

                border3.Append(leftBorder3);
                border3.Append(rightBorder3);
                border3.Append(topBorder3);
                border3.Append(bottomBorder3);
                border3.Append(diagonalBorder3);

                Border border4 = new Border();
                LeftBorder leftBorder4 = new LeftBorder();
                RightBorder rightBorder4 = new RightBorder();
                TopBorder topBorder4 = new TopBorder();

                BottomBorder bottomBorder4 = new BottomBorder() { Style = BorderStyleValues.Medium };
                Color color21 = new Color() { Theme = (UInt32Value)4U, Tint = 0.39997558519241921D };

                bottomBorder4.Append(color21);
                DiagonalBorder diagonalBorder4 = new DiagonalBorder();

                border4.Append(leftBorder4);
                border4.Append(rightBorder4);
                border4.Append(topBorder4);
                border4.Append(bottomBorder4);
                border4.Append(diagonalBorder4);

                Border border5 = new Border();

                LeftBorder leftBorder5 = new LeftBorder() { Style = BorderStyleValues.Thin };
                Color color22 = new Color() { Rgb = "FF7F7F7F" };

                leftBorder5.Append(color22);

                RightBorder rightBorder5 = new RightBorder() { Style = BorderStyleValues.Thin };
                Color color23 = new Color() { Rgb = "FF7F7F7F" };

                rightBorder5.Append(color23);

                TopBorder topBorder5 = new TopBorder() { Style = BorderStyleValues.Thin };
                Color color24 = new Color() { Rgb = "FF7F7F7F" };

                topBorder5.Append(color24);

                BottomBorder bottomBorder5 = new BottomBorder() { Style = BorderStyleValues.Thin };
                Color color25 = new Color() { Rgb = "FF7F7F7F" };

                bottomBorder5.Append(color25);
                DiagonalBorder diagonalBorder5 = new DiagonalBorder();

                border5.Append(leftBorder5);
                border5.Append(rightBorder5);
                border5.Append(topBorder5);
                border5.Append(bottomBorder5);
                border5.Append(diagonalBorder5);

                Border border6 = new Border();

                LeftBorder leftBorder6 = new LeftBorder() { Style = BorderStyleValues.Thin };
                Color color26 = new Color() { Rgb = "FF3F3F3F" };

                leftBorder6.Append(color26);

                RightBorder rightBorder6 = new RightBorder() { Style = BorderStyleValues.Thin };
                Color color27 = new Color() { Rgb = "FF3F3F3F" };

                rightBorder6.Append(color27);

                TopBorder topBorder6 = new TopBorder() { Style = BorderStyleValues.Thin };
                Color color28 = new Color() { Rgb = "FF3F3F3F" };

                topBorder6.Append(color28);

                BottomBorder bottomBorder6 = new BottomBorder() { Style = BorderStyleValues.Thin };
                Color color29 = new Color() { Rgb = "FF3F3F3F" };

                bottomBorder6.Append(color29);
                DiagonalBorder diagonalBorder6 = new DiagonalBorder();

                border6.Append(leftBorder6);
                border6.Append(rightBorder6);
                border6.Append(topBorder6);
                border6.Append(bottomBorder6);
                border6.Append(diagonalBorder6);

                Border border7 = new Border();
                LeftBorder leftBorder7 = new LeftBorder();
                RightBorder rightBorder7 = new RightBorder();
                TopBorder topBorder7 = new TopBorder();

                BottomBorder bottomBorder7 = new BottomBorder() { Style = BorderStyleValues.Double };
                Color color30 = new Color() { Rgb = "FFFF8001" };

                bottomBorder7.Append(color30);
                DiagonalBorder diagonalBorder7 = new DiagonalBorder();

                border7.Append(leftBorder7);
                border7.Append(rightBorder7);
                border7.Append(topBorder7);
                border7.Append(bottomBorder7);
                border7.Append(diagonalBorder7);

                Border border8 = new Border();

                LeftBorder leftBorder8 = new LeftBorder() { Style = BorderStyleValues.Double };
                Color color31 = new Color() { Rgb = "FF3F3F3F" };

                leftBorder8.Append(color31);

                RightBorder rightBorder8 = new RightBorder() { Style = BorderStyleValues.Double };
                Color color32 = new Color() { Rgb = "FF3F3F3F" };

                rightBorder8.Append(color32);

                TopBorder topBorder8 = new TopBorder() { Style = BorderStyleValues.Double };
                Color color33 = new Color() { Rgb = "FF3F3F3F" };

                topBorder8.Append(color33);

                BottomBorder bottomBorder8 = new BottomBorder() { Style = BorderStyleValues.Double };
                Color color34 = new Color() { Rgb = "FF3F3F3F" };

                bottomBorder8.Append(color34);
                DiagonalBorder diagonalBorder8 = new DiagonalBorder();

                border8.Append(leftBorder8);
                border8.Append(rightBorder8);
                border8.Append(topBorder8);
                border8.Append(bottomBorder8);
                border8.Append(diagonalBorder8);

                Border border9 = new Border();

                LeftBorder leftBorder9 = new LeftBorder() { Style = BorderStyleValues.Thin };
                Color color35 = new Color() { Rgb = "FFB2B2B2" };

                leftBorder9.Append(color35);

                RightBorder rightBorder9 = new RightBorder() { Style = BorderStyleValues.Thin };
                Color color36 = new Color() { Rgb = "FFB2B2B2" };

                rightBorder9.Append(color36);

                TopBorder topBorder9 = new TopBorder() { Style = BorderStyleValues.Thin };
                Color color37 = new Color() { Rgb = "FFB2B2B2" };

                topBorder9.Append(color37);

                BottomBorder bottomBorder9 = new BottomBorder() { Style = BorderStyleValues.Thin };
                Color color38 = new Color() { Rgb = "FFB2B2B2" };

                bottomBorder9.Append(color38);
                DiagonalBorder diagonalBorder9 = new DiagonalBorder();

                border9.Append(leftBorder9);
                border9.Append(rightBorder9);
                border9.Append(topBorder9);
                border9.Append(bottomBorder9);
                border9.Append(diagonalBorder9);

                Border border10 = new Border();
                LeftBorder leftBorder10 = new LeftBorder();
                RightBorder rightBorder10 = new RightBorder();

                TopBorder topBorder10 = new TopBorder() { Style = BorderStyleValues.Thin };
                Color color39 = new Color() { Theme = (UInt32Value)4U };

                topBorder10.Append(color39);

                BottomBorder bottomBorder10 = new BottomBorder() { Style = BorderStyleValues.Double };
                Color color40 = new Color() { Theme = (UInt32Value)4U };

                bottomBorder10.Append(color40);
                DiagonalBorder diagonalBorder10 = new DiagonalBorder();

                border10.Append(leftBorder10);
                border10.Append(rightBorder10);
                border10.Append(topBorder10);
                border10.Append(bottomBorder10);
                border10.Append(diagonalBorder10);

                borders1.Append(border1);
                borders1.Append(border2);
                borders1.Append(border3);
                borders1.Append(border4);
                borders1.Append(border5);
                borders1.Append(border6);
                borders1.Append(border7);
                borders1.Append(border8);
                borders1.Append(border9);
                borders1.Append(border10);

                CellStyleFormats cellStyleFormats1 = new CellStyleFormats() { Count = (UInt32Value)42U };
                CellFormat cellFormat1 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U };
                CellFormat cellFormat2 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)2U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, ApplyNumberFormat = false, ApplyFill = false, ApplyBorder = false, ApplyAlignment = false, ApplyProtection = false };
                CellFormat cellFormat3 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)3U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)1U, ApplyNumberFormat = false, ApplyFill = false, ApplyAlignment = false, ApplyProtection = false };
                CellFormat cellFormat4 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)4U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)2U, ApplyNumberFormat = false, ApplyFill = false, ApplyAlignment = false, ApplyProtection = false };
                CellFormat cellFormat5 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)5U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)3U, ApplyNumberFormat = false, ApplyFill = false, ApplyAlignment = false, ApplyProtection = false };
                CellFormat cellFormat6 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)5U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, ApplyNumberFormat = false, ApplyFill = false, ApplyBorder = false, ApplyAlignment = false, ApplyProtection = false };
                CellFormat cellFormat7 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)6U, FillId = (UInt32Value)2U, BorderId = (UInt32Value)0U, ApplyNumberFormat = false, ApplyBorder = false, ApplyAlignment = false, ApplyProtection = false };
                CellFormat cellFormat8 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)7U, FillId = (UInt32Value)3U, BorderId = (UInt32Value)0U, ApplyNumberFormat = false, ApplyBorder = false, ApplyAlignment = false, ApplyProtection = false };
                CellFormat cellFormat9 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)8U, FillId = (UInt32Value)4U, BorderId = (UInt32Value)0U, ApplyNumberFormat = false, ApplyBorder = false, ApplyAlignment = false, ApplyProtection = false };
                CellFormat cellFormat10 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)9U, FillId = (UInt32Value)5U, BorderId = (UInt32Value)4U, ApplyNumberFormat = false, ApplyAlignment = false, ApplyProtection = false };
                CellFormat cellFormat11 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)10U, FillId = (UInt32Value)6U, BorderId = (UInt32Value)5U, ApplyNumberFormat = false, ApplyAlignment = false, ApplyProtection = false };
                CellFormat cellFormat12 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)11U, FillId = (UInt32Value)6U, BorderId = (UInt32Value)4U, ApplyNumberFormat = false, ApplyAlignment = false, ApplyProtection = false };
                CellFormat cellFormat13 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)12U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)6U, ApplyNumberFormat = false, ApplyFill = false, ApplyAlignment = false, ApplyProtection = false };
                CellFormat cellFormat14 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)13U, FillId = (UInt32Value)7U, BorderId = (UInt32Value)7U, ApplyNumberFormat = false, ApplyAlignment = false, ApplyProtection = false };
                CellFormat cellFormat15 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)14U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, ApplyNumberFormat = false, ApplyFill = false, ApplyBorder = false, ApplyAlignment = false, ApplyProtection = false };
                CellFormat cellFormat16 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)1U, FillId = (UInt32Value)8U, BorderId = (UInt32Value)8U, ApplyNumberFormat = false, ApplyFont = false, ApplyAlignment = false, ApplyProtection = false };
                CellFormat cellFormat17 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)15U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, ApplyNumberFormat = false, ApplyFill = false, ApplyBorder = false, ApplyAlignment = false, ApplyProtection = false };
                CellFormat cellFormat18 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)16U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)9U, ApplyNumberFormat = false, ApplyFill = false, ApplyAlignment = false, ApplyProtection = false };
                CellFormat cellFormat19 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)17U, FillId = (UInt32Value)9U, BorderId = (UInt32Value)0U, ApplyNumberFormat = false, ApplyBorder = false, ApplyAlignment = false, ApplyProtection = false };
                CellFormat cellFormat20 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)1U, FillId = (UInt32Value)10U, BorderId = (UInt32Value)0U, ApplyNumberFormat = false, ApplyBorder = false, ApplyAlignment = false, ApplyProtection = false };
                CellFormat cellFormat21 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)1U, FillId = (UInt32Value)11U, BorderId = (UInt32Value)0U, ApplyNumberFormat = false, ApplyBorder = false, ApplyAlignment = false, ApplyProtection = false };
                CellFormat cellFormat22 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)17U, FillId = (UInt32Value)12U, BorderId = (UInt32Value)0U, ApplyNumberFormat = false, ApplyBorder = false, ApplyAlignment = false, ApplyProtection = false };
                CellFormat cellFormat23 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)17U, FillId = (UInt32Value)13U, BorderId = (UInt32Value)0U, ApplyNumberFormat = false, ApplyBorder = false, ApplyAlignment = false, ApplyProtection = false };
                CellFormat cellFormat24 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)1U, FillId = (UInt32Value)14U, BorderId = (UInt32Value)0U, ApplyNumberFormat = false, ApplyBorder = false, ApplyAlignment = false, ApplyProtection = false };
                CellFormat cellFormat25 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)1U, FillId = (UInt32Value)15U, BorderId = (UInt32Value)0U, ApplyNumberFormat = false, ApplyBorder = false, ApplyAlignment = false, ApplyProtection = false };
                CellFormat cellFormat26 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)17U, FillId = (UInt32Value)16U, BorderId = (UInt32Value)0U, ApplyNumberFormat = false, ApplyBorder = false, ApplyAlignment = false, ApplyProtection = false };
                CellFormat cellFormat27 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)17U, FillId = (UInt32Value)17U, BorderId = (UInt32Value)0U, ApplyNumberFormat = false, ApplyBorder = false, ApplyAlignment = false, ApplyProtection = false };
                CellFormat cellFormat28 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)1U, FillId = (UInt32Value)18U, BorderId = (UInt32Value)0U, ApplyNumberFormat = false, ApplyBorder = false, ApplyAlignment = false, ApplyProtection = false };
                CellFormat cellFormat29 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)1U, FillId = (UInt32Value)19U, BorderId = (UInt32Value)0U, ApplyNumberFormat = false, ApplyBorder = false, ApplyAlignment = false, ApplyProtection = false };
                CellFormat cellFormat30 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)17U, FillId = (UInt32Value)20U, BorderId = (UInt32Value)0U, ApplyNumberFormat = false, ApplyBorder = false, ApplyAlignment = false, ApplyProtection = false };
                CellFormat cellFormat31 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)17U, FillId = (UInt32Value)21U, BorderId = (UInt32Value)0U, ApplyNumberFormat = false, ApplyBorder = false, ApplyAlignment = false, ApplyProtection = false };
                CellFormat cellFormat32 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)1U, FillId = (UInt32Value)22U, BorderId = (UInt32Value)0U, ApplyNumberFormat = false, ApplyBorder = false, ApplyAlignment = false, ApplyProtection = false };
                CellFormat cellFormat33 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)1U, FillId = (UInt32Value)23U, BorderId = (UInt32Value)0U, ApplyNumberFormat = false, ApplyBorder = false, ApplyAlignment = false, ApplyProtection = false };
                CellFormat cellFormat34 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)17U, FillId = (UInt32Value)24U, BorderId = (UInt32Value)0U, ApplyNumberFormat = false, ApplyBorder = false, ApplyAlignment = false, ApplyProtection = false };
                CellFormat cellFormat35 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)17U, FillId = (UInt32Value)25U, BorderId = (UInt32Value)0U, ApplyNumberFormat = false, ApplyBorder = false, ApplyAlignment = false, ApplyProtection = false };
                CellFormat cellFormat36 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)1U, FillId = (UInt32Value)26U, BorderId = (UInt32Value)0U, ApplyNumberFormat = false, ApplyBorder = false, ApplyAlignment = false, ApplyProtection = false };
                CellFormat cellFormat37 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)1U, FillId = (UInt32Value)27U, BorderId = (UInt32Value)0U, ApplyNumberFormat = false, ApplyBorder = false, ApplyAlignment = false, ApplyProtection = false };
                CellFormat cellFormat38 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)17U, FillId = (UInt32Value)28U, BorderId = (UInt32Value)0U, ApplyNumberFormat = false, ApplyBorder = false, ApplyAlignment = false, ApplyProtection = false };
                CellFormat cellFormat39 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)17U, FillId = (UInt32Value)29U, BorderId = (UInt32Value)0U, ApplyNumberFormat = false, ApplyBorder = false, ApplyAlignment = false, ApplyProtection = false };
                CellFormat cellFormat40 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)1U, FillId = (UInt32Value)30U, BorderId = (UInt32Value)0U, ApplyNumberFormat = false, ApplyBorder = false, ApplyAlignment = false, ApplyProtection = false };
                CellFormat cellFormat41 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)1U, FillId = (UInt32Value)31U, BorderId = (UInt32Value)0U, ApplyNumberFormat = false, ApplyBorder = false, ApplyAlignment = false, ApplyProtection = false };
                CellFormat cellFormat42 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)17U, FillId = (UInt32Value)32U, BorderId = (UInt32Value)0U, ApplyNumberFormat = false, ApplyBorder = false, ApplyAlignment = false, ApplyProtection = false };

                cellStyleFormats1.Append(cellFormat1);
                cellStyleFormats1.Append(cellFormat2);
                cellStyleFormats1.Append(cellFormat3);
                cellStyleFormats1.Append(cellFormat4);
                cellStyleFormats1.Append(cellFormat5);
                cellStyleFormats1.Append(cellFormat6);
                cellStyleFormats1.Append(cellFormat7);
                cellStyleFormats1.Append(cellFormat8);
                cellStyleFormats1.Append(cellFormat9);
                cellStyleFormats1.Append(cellFormat10);
                cellStyleFormats1.Append(cellFormat11);
                cellStyleFormats1.Append(cellFormat12);
                cellStyleFormats1.Append(cellFormat13);
                cellStyleFormats1.Append(cellFormat14);
                cellStyleFormats1.Append(cellFormat15);
                cellStyleFormats1.Append(cellFormat16);
                cellStyleFormats1.Append(cellFormat17);
                cellStyleFormats1.Append(cellFormat18);
                cellStyleFormats1.Append(cellFormat19);
                cellStyleFormats1.Append(cellFormat20);
                cellStyleFormats1.Append(cellFormat21);
                cellStyleFormats1.Append(cellFormat22);
                cellStyleFormats1.Append(cellFormat23);
                cellStyleFormats1.Append(cellFormat24);
                cellStyleFormats1.Append(cellFormat25);
                cellStyleFormats1.Append(cellFormat26);
                cellStyleFormats1.Append(cellFormat27);
                cellStyleFormats1.Append(cellFormat28);
                cellStyleFormats1.Append(cellFormat29);
                cellStyleFormats1.Append(cellFormat30);
                cellStyleFormats1.Append(cellFormat31);
                cellStyleFormats1.Append(cellFormat32);
                cellStyleFormats1.Append(cellFormat33);
                cellStyleFormats1.Append(cellFormat34);
                cellStyleFormats1.Append(cellFormat35);
                cellStyleFormats1.Append(cellFormat36);
                cellStyleFormats1.Append(cellFormat37);
                cellStyleFormats1.Append(cellFormat38);
                cellStyleFormats1.Append(cellFormat39);
                cellStyleFormats1.Append(cellFormat40);
                cellStyleFormats1.Append(cellFormat41);
                cellStyleFormats1.Append(cellFormat42);

                CellFormats cellFormats1 = new CellFormats() { Count = (UInt32Value)8U };
                CellFormat cellFormat43 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U };
                CellFormat cellFormat44 = new CellFormat() { NumberFormatId = (UInt32Value)49U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U, ApplyNumberFormat = true };
                CellFormat cellFormat45 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U, PivotButton = true };

                CellFormat cellFormat46 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U, ApplyAlignment = true };
                Alignment alignment1 = new Alignment() { Horizontal = HorizontalAlignmentValues.Left };

                cellFormat46.Append(alignment1);

                CellFormat cellFormat47 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U, ApplyAlignment = true };
                Alignment alignment2 = new Alignment() { Horizontal = HorizontalAlignmentValues.Left, Indent = (UInt32Value)1U };

                cellFormat47.Append(alignment2);

                CellFormat cellFormat48 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U, ApplyAlignment = true };
                Alignment alignment3 = new Alignment() { Horizontal = HorizontalAlignmentValues.Left, Indent = (UInt32Value)2U };

                cellFormat48.Append(alignment3);

                CellFormat cellFormat49 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U, ApplyAlignment = true };
                Alignment alignment4 = new Alignment() { Horizontal = HorizontalAlignmentValues.Left, Indent = (UInt32Value)3U };

                cellFormat49.Append(alignment4);
                CellFormat cellFormat50 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U, ApplyNumberFormat = true };

                cellFormats1.Append(cellFormat43);
                cellFormats1.Append(cellFormat44);
                cellFormats1.Append(cellFormat45);
                cellFormats1.Append(cellFormat46);
                cellFormats1.Append(cellFormat47);
                cellFormats1.Append(cellFormat48);
                cellFormats1.Append(cellFormat49);
                cellFormats1.Append(cellFormat50);

                CellStyles cellStyles1 = new CellStyles() { Count = (UInt32Value)42U };
                CellStyle cellStyle1 = new CellStyle() { Name = "20% — акцент1", FormatId = (UInt32Value)19U, BuiltinId = (UInt32Value)30U, CustomBuiltin = true };
                CellStyle cellStyle2 = new CellStyle() { Name = "20% — акцент2", FormatId = (UInt32Value)23U, BuiltinId = (UInt32Value)34U, CustomBuiltin = true };
                CellStyle cellStyle3 = new CellStyle() { Name = "20% — акцент3", FormatId = (UInt32Value)27U, BuiltinId = (UInt32Value)38U, CustomBuiltin = true };
                CellStyle cellStyle4 = new CellStyle() { Name = "20% — акцент4", FormatId = (UInt32Value)31U, BuiltinId = (UInt32Value)42U, CustomBuiltin = true };
                CellStyle cellStyle5 = new CellStyle() { Name = "20% — акцент5", FormatId = (UInt32Value)35U, BuiltinId = (UInt32Value)46U, CustomBuiltin = true };
                CellStyle cellStyle6 = new CellStyle() { Name = "20% — акцент6", FormatId = (UInt32Value)39U, BuiltinId = (UInt32Value)50U, CustomBuiltin = true };
                CellStyle cellStyle7 = new CellStyle() { Name = "40% — акцент1", FormatId = (UInt32Value)20U, BuiltinId = (UInt32Value)31U, CustomBuiltin = true };
                CellStyle cellStyle8 = new CellStyle() { Name = "40% — акцент2", FormatId = (UInt32Value)24U, BuiltinId = (UInt32Value)35U, CustomBuiltin = true };
                CellStyle cellStyle9 = new CellStyle() { Name = "40% — акцент3", FormatId = (UInt32Value)28U, BuiltinId = (UInt32Value)39U, CustomBuiltin = true };
                CellStyle cellStyle10 = new CellStyle() { Name = "40% — акцент4", FormatId = (UInt32Value)32U, BuiltinId = (UInt32Value)43U, CustomBuiltin = true };
                CellStyle cellStyle11 = new CellStyle() { Name = "40% — акцент5", FormatId = (UInt32Value)36U, BuiltinId = (UInt32Value)47U, CustomBuiltin = true };
                CellStyle cellStyle12 = new CellStyle() { Name = "40% — акцент6", FormatId = (UInt32Value)40U, BuiltinId = (UInt32Value)51U, CustomBuiltin = true };
                CellStyle cellStyle13 = new CellStyle() { Name = "60% — акцент1", FormatId = (UInt32Value)21U, BuiltinId = (UInt32Value)32U, CustomBuiltin = true };
                CellStyle cellStyle14 = new CellStyle() { Name = "60% — акцент2", FormatId = (UInt32Value)25U, BuiltinId = (UInt32Value)36U, CustomBuiltin = true };
                CellStyle cellStyle15 = new CellStyle() { Name = "60% — акцент3", FormatId = (UInt32Value)29U, BuiltinId = (UInt32Value)40U, CustomBuiltin = true };
                CellStyle cellStyle16 = new CellStyle() { Name = "60% — акцент4", FormatId = (UInt32Value)33U, BuiltinId = (UInt32Value)44U, CustomBuiltin = true };
                CellStyle cellStyle17 = new CellStyle() { Name = "60% — акцент5", FormatId = (UInt32Value)37U, BuiltinId = (UInt32Value)48U, CustomBuiltin = true };
                CellStyle cellStyle18 = new CellStyle() { Name = "60% — акцент6", FormatId = (UInt32Value)41U, BuiltinId = (UInt32Value)52U, CustomBuiltin = true };
                CellStyle cellStyle19 = new CellStyle() { Name = "Акцент1", FormatId = (UInt32Value)18U, BuiltinId = (UInt32Value)29U, CustomBuiltin = true };
                CellStyle cellStyle20 = new CellStyle() { Name = "Акцент2", FormatId = (UInt32Value)22U, BuiltinId = (UInt32Value)33U, CustomBuiltin = true };
                CellStyle cellStyle21 = new CellStyle() { Name = "Акцент3", FormatId = (UInt32Value)26U, BuiltinId = (UInt32Value)37U, CustomBuiltin = true };
                CellStyle cellStyle22 = new CellStyle() { Name = "Акцент4", FormatId = (UInt32Value)30U, BuiltinId = (UInt32Value)41U, CustomBuiltin = true };
                CellStyle cellStyle23 = new CellStyle() { Name = "Акцент5", FormatId = (UInt32Value)34U, BuiltinId = (UInt32Value)45U, CustomBuiltin = true };
                CellStyle cellStyle24 = new CellStyle() { Name = "Акцент6", FormatId = (UInt32Value)38U, BuiltinId = (UInt32Value)49U, CustomBuiltin = true };
                CellStyle cellStyle25 = new CellStyle() { Name = "Ввод ", FormatId = (UInt32Value)9U, BuiltinId = (UInt32Value)20U, CustomBuiltin = true };
                CellStyle cellStyle26 = new CellStyle() { Name = "Вывод", FormatId = (UInt32Value)10U, BuiltinId = (UInt32Value)21U, CustomBuiltin = true };
                CellStyle cellStyle27 = new CellStyle() { Name = "Вычисление", FormatId = (UInt32Value)11U, BuiltinId = (UInt32Value)22U, CustomBuiltin = true };
                CellStyle cellStyle28 = new CellStyle() { Name = "Заголовок 1", FormatId = (UInt32Value)2U, BuiltinId = (UInt32Value)16U, CustomBuiltin = true };
                CellStyle cellStyle29 = new CellStyle() { Name = "Заголовок 2", FormatId = (UInt32Value)3U, BuiltinId = (UInt32Value)17U, CustomBuiltin = true };
                CellStyle cellStyle30 = new CellStyle() { Name = "Заголовок 3", FormatId = (UInt32Value)4U, BuiltinId = (UInt32Value)18U, CustomBuiltin = true };
                CellStyle cellStyle31 = new CellStyle() { Name = "Заголовок 4", FormatId = (UInt32Value)5U, BuiltinId = (UInt32Value)19U, CustomBuiltin = true };
                CellStyle cellStyle32 = new CellStyle() { Name = "Итог", FormatId = (UInt32Value)17U, BuiltinId = (UInt32Value)25U, CustomBuiltin = true };
                CellStyle cellStyle33 = new CellStyle() { Name = "Контрольная ячейка", FormatId = (UInt32Value)13U, BuiltinId = (UInt32Value)23U, CustomBuiltin = true };
                CellStyle cellStyle34 = new CellStyle() { Name = "Название", FormatId = (UInt32Value)1U, BuiltinId = (UInt32Value)15U, CustomBuiltin = true };
                CellStyle cellStyle35 = new CellStyle() { Name = "Нейтральный", FormatId = (UInt32Value)8U, BuiltinId = (UInt32Value)28U, CustomBuiltin = true };
                CellStyle cellStyle36 = new CellStyle() { Name = "Обычный", FormatId = (UInt32Value)0U, BuiltinId = (UInt32Value)0U };
                CellStyle cellStyle37 = new CellStyle() { Name = "Плохой", FormatId = (UInt32Value)7U, BuiltinId = (UInt32Value)27U, CustomBuiltin = true };
                CellStyle cellStyle38 = new CellStyle() { Name = "Пояснение", FormatId = (UInt32Value)16U, BuiltinId = (UInt32Value)53U, CustomBuiltin = true };
                CellStyle cellStyle39 = new CellStyle() { Name = "Примечание", FormatId = (UInt32Value)15U, BuiltinId = (UInt32Value)10U, CustomBuiltin = true };
                CellStyle cellStyle40 = new CellStyle() { Name = "Связанная ячейка", FormatId = (UInt32Value)12U, BuiltinId = (UInt32Value)24U, CustomBuiltin = true };
                CellStyle cellStyle41 = new CellStyle() { Name = "Текст предупреждения", FormatId = (UInt32Value)14U, BuiltinId = (UInt32Value)11U, CustomBuiltin = true };
                CellStyle cellStyle42 = new CellStyle() { Name = "Хороший", FormatId = (UInt32Value)6U, BuiltinId = (UInt32Value)26U, CustomBuiltin = true };

                cellStyles1.Append(cellStyle1);
                cellStyles1.Append(cellStyle2);
                cellStyles1.Append(cellStyle3);
                cellStyles1.Append(cellStyle4);
                cellStyles1.Append(cellStyle5);
                cellStyles1.Append(cellStyle6);
                cellStyles1.Append(cellStyle7);
                cellStyles1.Append(cellStyle8);
                cellStyles1.Append(cellStyle9);
                cellStyles1.Append(cellStyle10);
                cellStyles1.Append(cellStyle11);
                cellStyles1.Append(cellStyle12);
                cellStyles1.Append(cellStyle13);
                cellStyles1.Append(cellStyle14);
                cellStyles1.Append(cellStyle15);
                cellStyles1.Append(cellStyle16);
                cellStyles1.Append(cellStyle17);
                cellStyles1.Append(cellStyle18);
                cellStyles1.Append(cellStyle19);
                cellStyles1.Append(cellStyle20);
                cellStyles1.Append(cellStyle21);
                cellStyles1.Append(cellStyle22);
                cellStyles1.Append(cellStyle23);
                cellStyles1.Append(cellStyle24);
                cellStyles1.Append(cellStyle25);
                cellStyles1.Append(cellStyle26);
                cellStyles1.Append(cellStyle27);
                cellStyles1.Append(cellStyle28);
                cellStyles1.Append(cellStyle29);
                cellStyles1.Append(cellStyle30);
                cellStyles1.Append(cellStyle31);
                cellStyles1.Append(cellStyle32);
                cellStyles1.Append(cellStyle33);
                cellStyles1.Append(cellStyle34);
                cellStyles1.Append(cellStyle35);
                cellStyles1.Append(cellStyle36);
                cellStyles1.Append(cellStyle37);
                cellStyles1.Append(cellStyle38);
                cellStyles1.Append(cellStyle39);
                cellStyles1.Append(cellStyle40);
                cellStyles1.Append(cellStyle41);
                cellStyles1.Append(cellStyle42);

                DifferentialFormats differentialFormats1 = new DifferentialFormats() { Count = (UInt32Value)5U };

                DifferentialFormat differentialFormat1 = new DifferentialFormat();
                NumberingFormat numberingFormat1 = new NumberingFormat() { NumberFormatId = (UInt32Value)0U, FormatCode = "General" };

                differentialFormat1.Append(numberingFormat1);

                DifferentialFormat differentialFormat2 = new DifferentialFormat();
                NumberingFormat numberingFormat2 = new NumberingFormat() { NumberFormatId = (UInt32Value)30U, FormatCode = "@" };

                differentialFormat2.Append(numberingFormat2);

                DifferentialFormat differentialFormat3 = new DifferentialFormat();
                NumberingFormat numberingFormat3 = new NumberingFormat() { NumberFormatId = (UInt32Value)30U, FormatCode = "@" };

                differentialFormat3.Append(numberingFormat3);

                DifferentialFormat differentialFormat4 = new DifferentialFormat();
                NumberingFormat numberingFormat4 = new NumberingFormat() { NumberFormatId = (UInt32Value)30U, FormatCode = "@" };

                differentialFormat4.Append(numberingFormat4);

                DifferentialFormat differentialFormat5 = new DifferentialFormat();
                NumberingFormat numberingFormat5 = new NumberingFormat() { NumberFormatId = (UInt32Value)30U, FormatCode = "@" };

                differentialFormat5.Append(numberingFormat5);

                differentialFormats1.Append(differentialFormat1);
                differentialFormats1.Append(differentialFormat2);
                differentialFormats1.Append(differentialFormat3);
                differentialFormats1.Append(differentialFormat4);
                differentialFormats1.Append(differentialFormat5);
                TableStyles tableStyles1 = new TableStyles() { Count = (UInt32Value)0U, DefaultTableStyle = "TableStyleMedium2", DefaultPivotStyle = "PivotStyleLight16" };

                StylesheetExtensionList stylesheetExtensionList1 = new StylesheetExtensionList();

                StylesheetExtension stylesheetExtension1 = new StylesheetExtension() { Uri = "{EB79DEF2-80B8-43e5-95BD-54CBDDF9020C}" };
                stylesheetExtension1.AddNamespaceDeclaration("x14", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main");
                X14.SlicerStyles slicerStyles1 = new X14.SlicerStyles() { DefaultSlicerStyle = "SlicerStyleLight1" };

                stylesheetExtension1.Append(slicerStyles1);

                StylesheetExtension stylesheetExtension2 = new StylesheetExtension() { Uri = "{9260A510-F301-46a8-8635-F512D64BE5F5}" };
                stylesheetExtension2.AddNamespaceDeclaration("x15", "http://schemas.microsoft.com/office/spreadsheetml/2010/11/main");

                OpenXmlUnknownElement openXmlUnknownElement2 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<x15:timelineStyles defaultTimelineStyle=\"TimeSlicerStyleLight1\" xmlns:x15=\"http://schemas.microsoft.com/office/spreadsheetml/2010/11/main\" />");

                stylesheetExtension2.Append(openXmlUnknownElement2);

                stylesheetExtensionList1.Append(stylesheetExtension1);
                stylesheetExtensionList1.Append(stylesheetExtension2);

                stylesheet1.Append(fonts1);
                stylesheet1.Append(fills1);
                stylesheet1.Append(borders1);
                stylesheet1.Append(cellStyleFormats1);
                stylesheet1.Append(cellFormats1);
                stylesheet1.Append(cellStyles1);
                stylesheet1.Append(differentialFormats1);
                stylesheet1.Append(tableStyles1);
                stylesheet1.Append(stylesheetExtensionList1);

                workbookStylesPart1.Stylesheet = stylesheet1;
            }

            // Generates content of themePart1.
            private void GenerateThemePart1Content(ThemePart themePart1)
            {
                A.Theme theme1 = new A.Theme() { Name = "Тема Office" };
                theme1.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");

                A.ThemeElements themeElements1 = new A.ThemeElements();

                A.ColorScheme colorScheme1 = new A.ColorScheme() { Name = "Стандартная" };

                A.Dark1Color dark1Color1 = new A.Dark1Color();
                A.SystemColor systemColor1 = new A.SystemColor() { Val = A.SystemColorValues.WindowText, LastColor = "000000" };

                dark1Color1.Append(systemColor1);

                A.Light1Color light1Color1 = new A.Light1Color();
                A.SystemColor systemColor2 = new A.SystemColor() { Val = A.SystemColorValues.Window, LastColor = "FFFFFF" };

                light1Color1.Append(systemColor2);

                A.Dark2Color dark2Color1 = new A.Dark2Color();
                A.RgbColorModelHex rgbColorModelHex1 = new A.RgbColorModelHex() { Val = "44546A" };

                dark2Color1.Append(rgbColorModelHex1);

                A.Light2Color light2Color1 = new A.Light2Color();
                A.RgbColorModelHex rgbColorModelHex2 = new A.RgbColorModelHex() { Val = "E7E6E6" };

                light2Color1.Append(rgbColorModelHex2);

                A.Accent1Color accent1Color1 = new A.Accent1Color();
                A.RgbColorModelHex rgbColorModelHex3 = new A.RgbColorModelHex() { Val = "5B9BD5" };

                accent1Color1.Append(rgbColorModelHex3);

                A.Accent2Color accent2Color1 = new A.Accent2Color();
                A.RgbColorModelHex rgbColorModelHex4 = new A.RgbColorModelHex() { Val = "ED7D31" };

                accent2Color1.Append(rgbColorModelHex4);

                A.Accent3Color accent3Color1 = new A.Accent3Color();
                A.RgbColorModelHex rgbColorModelHex5 = new A.RgbColorModelHex() { Val = "A5A5A5" };

                accent3Color1.Append(rgbColorModelHex5);

                A.Accent4Color accent4Color1 = new A.Accent4Color();
                A.RgbColorModelHex rgbColorModelHex6 = new A.RgbColorModelHex() { Val = "FFC000" };

                accent4Color1.Append(rgbColorModelHex6);

                A.Accent5Color accent5Color1 = new A.Accent5Color();
                A.RgbColorModelHex rgbColorModelHex7 = new A.RgbColorModelHex() { Val = "4472C4" };

                accent5Color1.Append(rgbColorModelHex7);

                A.Accent6Color accent6Color1 = new A.Accent6Color();
                A.RgbColorModelHex rgbColorModelHex8 = new A.RgbColorModelHex() { Val = "70AD47" };

                accent6Color1.Append(rgbColorModelHex8);

                A.Hyperlink hyperlink1 = new A.Hyperlink();
                A.RgbColorModelHex rgbColorModelHex9 = new A.RgbColorModelHex() { Val = "0563C1" };

                hyperlink1.Append(rgbColorModelHex9);

                A.FollowedHyperlinkColor followedHyperlinkColor1 = new A.FollowedHyperlinkColor();
                A.RgbColorModelHex rgbColorModelHex10 = new A.RgbColorModelHex() { Val = "954F72" };

                followedHyperlinkColor1.Append(rgbColorModelHex10);

                colorScheme1.Append(dark1Color1);
                colorScheme1.Append(light1Color1);
                colorScheme1.Append(dark2Color1);
                colorScheme1.Append(light2Color1);
                colorScheme1.Append(accent1Color1);
                colorScheme1.Append(accent2Color1);
                colorScheme1.Append(accent3Color1);
                colorScheme1.Append(accent4Color1);
                colorScheme1.Append(accent5Color1);
                colorScheme1.Append(accent6Color1);
                colorScheme1.Append(hyperlink1);
                colorScheme1.Append(followedHyperlinkColor1);

                A.FontScheme fontScheme19 = new A.FontScheme() { Name = "Стандартная" };

                A.MajorFont majorFont1 = new A.MajorFont();
                A.LatinFont latinFont1 = new A.LatinFont() { Typeface = "Calibri Light", Panose = "020F0302020204030204" };
                A.EastAsianFont eastAsianFont1 = new A.EastAsianFont() { Typeface = "" };
                A.ComplexScriptFont complexScriptFont1 = new A.ComplexScriptFont() { Typeface = "" };
                A.SupplementalFont supplementalFont1 = new A.SupplementalFont() { Script = "Jpan", Typeface = "游ゴシック Light" };
                A.SupplementalFont supplementalFont2 = new A.SupplementalFont() { Script = "Hang", Typeface = "맑은 고딕" };
                A.SupplementalFont supplementalFont3 = new A.SupplementalFont() { Script = "Hans", Typeface = "等线 Light" };
                A.SupplementalFont supplementalFont4 = new A.SupplementalFont() { Script = "Hant", Typeface = "新細明體" };
                A.SupplementalFont supplementalFont5 = new A.SupplementalFont() { Script = "Arab", Typeface = "Times New Roman" };
                A.SupplementalFont supplementalFont6 = new A.SupplementalFont() { Script = "Hebr", Typeface = "Times New Roman" };
                A.SupplementalFont supplementalFont7 = new A.SupplementalFont() { Script = "Thai", Typeface = "Tahoma" };
                A.SupplementalFont supplementalFont8 = new A.SupplementalFont() { Script = "Ethi", Typeface = "Nyala" };
                A.SupplementalFont supplementalFont9 = new A.SupplementalFont() { Script = "Beng", Typeface = "Vrinda" };
                A.SupplementalFont supplementalFont10 = new A.SupplementalFont() { Script = "Gujr", Typeface = "Shruti" };
                A.SupplementalFont supplementalFont11 = new A.SupplementalFont() { Script = "Khmr", Typeface = "MoolBoran" };
                A.SupplementalFont supplementalFont12 = new A.SupplementalFont() { Script = "Knda", Typeface = "Tunga" };
                A.SupplementalFont supplementalFont13 = new A.SupplementalFont() { Script = "Guru", Typeface = "Raavi" };
                A.SupplementalFont supplementalFont14 = new A.SupplementalFont() { Script = "Cans", Typeface = "Euphemia" };
                A.SupplementalFont supplementalFont15 = new A.SupplementalFont() { Script = "Cher", Typeface = "Plantagenet Cherokee" };
                A.SupplementalFont supplementalFont16 = new A.SupplementalFont() { Script = "Yiii", Typeface = "Microsoft Yi Baiti" };
                A.SupplementalFont supplementalFont17 = new A.SupplementalFont() { Script = "Tibt", Typeface = "Microsoft Himalaya" };
                A.SupplementalFont supplementalFont18 = new A.SupplementalFont() { Script = "Thaa", Typeface = "MV Boli" };
                A.SupplementalFont supplementalFont19 = new A.SupplementalFont() { Script = "Deva", Typeface = "Mangal" };
                A.SupplementalFont supplementalFont20 = new A.SupplementalFont() { Script = "Telu", Typeface = "Gautami" };
                A.SupplementalFont supplementalFont21 = new A.SupplementalFont() { Script = "Taml", Typeface = "Latha" };
                A.SupplementalFont supplementalFont22 = new A.SupplementalFont() { Script = "Syrc", Typeface = "Estrangelo Edessa" };
                A.SupplementalFont supplementalFont23 = new A.SupplementalFont() { Script = "Orya", Typeface = "Kalinga" };
                A.SupplementalFont supplementalFont24 = new A.SupplementalFont() { Script = "Mlym", Typeface = "Kartika" };
                A.SupplementalFont supplementalFont25 = new A.SupplementalFont() { Script = "Laoo", Typeface = "DokChampa" };
                A.SupplementalFont supplementalFont26 = new A.SupplementalFont() { Script = "Sinh", Typeface = "Iskoola Pota" };
                A.SupplementalFont supplementalFont27 = new A.SupplementalFont() { Script = "Mong", Typeface = "Mongolian Baiti" };
                A.SupplementalFont supplementalFont28 = new A.SupplementalFont() { Script = "Viet", Typeface = "Times New Roman" };
                A.SupplementalFont supplementalFont29 = new A.SupplementalFont() { Script = "Uigh", Typeface = "Microsoft Uighur" };
                A.SupplementalFont supplementalFont30 = new A.SupplementalFont() { Script = "Geor", Typeface = "Sylfaen" };

                majorFont1.Append(latinFont1);
                majorFont1.Append(eastAsianFont1);
                majorFont1.Append(complexScriptFont1);
                majorFont1.Append(supplementalFont1);
                majorFont1.Append(supplementalFont2);
                majorFont1.Append(supplementalFont3);
                majorFont1.Append(supplementalFont4);
                majorFont1.Append(supplementalFont5);
                majorFont1.Append(supplementalFont6);
                majorFont1.Append(supplementalFont7);
                majorFont1.Append(supplementalFont8);
                majorFont1.Append(supplementalFont9);
                majorFont1.Append(supplementalFont10);
                majorFont1.Append(supplementalFont11);
                majorFont1.Append(supplementalFont12);
                majorFont1.Append(supplementalFont13);
                majorFont1.Append(supplementalFont14);
                majorFont1.Append(supplementalFont15);
                majorFont1.Append(supplementalFont16);
                majorFont1.Append(supplementalFont17);
                majorFont1.Append(supplementalFont18);
                majorFont1.Append(supplementalFont19);
                majorFont1.Append(supplementalFont20);
                majorFont1.Append(supplementalFont21);
                majorFont1.Append(supplementalFont22);
                majorFont1.Append(supplementalFont23);
                majorFont1.Append(supplementalFont24);
                majorFont1.Append(supplementalFont25);
                majorFont1.Append(supplementalFont26);
                majorFont1.Append(supplementalFont27);
                majorFont1.Append(supplementalFont28);
                majorFont1.Append(supplementalFont29);
                majorFont1.Append(supplementalFont30);

                A.MinorFont minorFont1 = new A.MinorFont();
                A.LatinFont latinFont2 = new A.LatinFont() { Typeface = "Calibri", Panose = "020F0502020204030204" };
                A.EastAsianFont eastAsianFont2 = new A.EastAsianFont() { Typeface = "" };
                A.ComplexScriptFont complexScriptFont2 = new A.ComplexScriptFont() { Typeface = "" };
                A.SupplementalFont supplementalFont31 = new A.SupplementalFont() { Script = "Jpan", Typeface = "游ゴシック" };
                A.SupplementalFont supplementalFont32 = new A.SupplementalFont() { Script = "Hang", Typeface = "맑은 고딕" };
                A.SupplementalFont supplementalFont33 = new A.SupplementalFont() { Script = "Hans", Typeface = "等线" };
                A.SupplementalFont supplementalFont34 = new A.SupplementalFont() { Script = "Hant", Typeface = "新細明體" };
                A.SupplementalFont supplementalFont35 = new A.SupplementalFont() { Script = "Arab", Typeface = "Arial" };
                A.SupplementalFont supplementalFont36 = new A.SupplementalFont() { Script = "Hebr", Typeface = "Arial" };
                A.SupplementalFont supplementalFont37 = new A.SupplementalFont() { Script = "Thai", Typeface = "Tahoma" };
                A.SupplementalFont supplementalFont38 = new A.SupplementalFont() { Script = "Ethi", Typeface = "Nyala" };
                A.SupplementalFont supplementalFont39 = new A.SupplementalFont() { Script = "Beng", Typeface = "Vrinda" };
                A.SupplementalFont supplementalFont40 = new A.SupplementalFont() { Script = "Gujr", Typeface = "Shruti" };
                A.SupplementalFont supplementalFont41 = new A.SupplementalFont() { Script = "Khmr", Typeface = "DaunPenh" };
                A.SupplementalFont supplementalFont42 = new A.SupplementalFont() { Script = "Knda", Typeface = "Tunga" };
                A.SupplementalFont supplementalFont43 = new A.SupplementalFont() { Script = "Guru", Typeface = "Raavi" };
                A.SupplementalFont supplementalFont44 = new A.SupplementalFont() { Script = "Cans", Typeface = "Euphemia" };
                A.SupplementalFont supplementalFont45 = new A.SupplementalFont() { Script = "Cher", Typeface = "Plantagenet Cherokee" };
                A.SupplementalFont supplementalFont46 = new A.SupplementalFont() { Script = "Yiii", Typeface = "Microsoft Yi Baiti" };
                A.SupplementalFont supplementalFont47 = new A.SupplementalFont() { Script = "Tibt", Typeface = "Microsoft Himalaya" };
                A.SupplementalFont supplementalFont48 = new A.SupplementalFont() { Script = "Thaa", Typeface = "MV Boli" };
                A.SupplementalFont supplementalFont49 = new A.SupplementalFont() { Script = "Deva", Typeface = "Mangal" };
                A.SupplementalFont supplementalFont50 = new A.SupplementalFont() { Script = "Telu", Typeface = "Gautami" };
                A.SupplementalFont supplementalFont51 = new A.SupplementalFont() { Script = "Taml", Typeface = "Latha" };
                A.SupplementalFont supplementalFont52 = new A.SupplementalFont() { Script = "Syrc", Typeface = "Estrangelo Edessa" };
                A.SupplementalFont supplementalFont53 = new A.SupplementalFont() { Script = "Orya", Typeface = "Kalinga" };
                A.SupplementalFont supplementalFont54 = new A.SupplementalFont() { Script = "Mlym", Typeface = "Kartika" };
                A.SupplementalFont supplementalFont55 = new A.SupplementalFont() { Script = "Laoo", Typeface = "DokChampa" };
                A.SupplementalFont supplementalFont56 = new A.SupplementalFont() { Script = "Sinh", Typeface = "Iskoola Pota" };
                A.SupplementalFont supplementalFont57 = new A.SupplementalFont() { Script = "Mong", Typeface = "Mongolian Baiti" };
                A.SupplementalFont supplementalFont58 = new A.SupplementalFont() { Script = "Viet", Typeface = "Arial" };
                A.SupplementalFont supplementalFont59 = new A.SupplementalFont() { Script = "Uigh", Typeface = "Microsoft Uighur" };
                A.SupplementalFont supplementalFont60 = new A.SupplementalFont() { Script = "Geor", Typeface = "Sylfaen" };

                minorFont1.Append(latinFont2);
                minorFont1.Append(eastAsianFont2);
                minorFont1.Append(complexScriptFont2);
                minorFont1.Append(supplementalFont31);
                minorFont1.Append(supplementalFont32);
                minorFont1.Append(supplementalFont33);
                minorFont1.Append(supplementalFont34);
                minorFont1.Append(supplementalFont35);
                minorFont1.Append(supplementalFont36);
                minorFont1.Append(supplementalFont37);
                minorFont1.Append(supplementalFont38);
                minorFont1.Append(supplementalFont39);
                minorFont1.Append(supplementalFont40);
                minorFont1.Append(supplementalFont41);
                minorFont1.Append(supplementalFont42);
                minorFont1.Append(supplementalFont43);
                minorFont1.Append(supplementalFont44);
                minorFont1.Append(supplementalFont45);
                minorFont1.Append(supplementalFont46);
                minorFont1.Append(supplementalFont47);
                minorFont1.Append(supplementalFont48);
                minorFont1.Append(supplementalFont49);
                minorFont1.Append(supplementalFont50);
                minorFont1.Append(supplementalFont51);
                minorFont1.Append(supplementalFont52);
                minorFont1.Append(supplementalFont53);
                minorFont1.Append(supplementalFont54);
                minorFont1.Append(supplementalFont55);
                minorFont1.Append(supplementalFont56);
                minorFont1.Append(supplementalFont57);
                minorFont1.Append(supplementalFont58);
                minorFont1.Append(supplementalFont59);
                minorFont1.Append(supplementalFont60);

                fontScheme19.Append(majorFont1);
                fontScheme19.Append(minorFont1);

                A.FormatScheme formatScheme1 = new A.FormatScheme() { Name = "Стандартная" };

                A.FillStyleList fillStyleList1 = new A.FillStyleList();

                A.SolidFill solidFill1 = new A.SolidFill();
                A.SchemeColor schemeColor1 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };

                solidFill1.Append(schemeColor1);

                A.GradientFill gradientFill1 = new A.GradientFill() { RotateWithShape = true };

                A.GradientStopList gradientStopList1 = new A.GradientStopList();

                A.GradientStop gradientStop1 = new A.GradientStop() { Position = 0 };

                A.SchemeColor schemeColor2 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
                A.LuminanceModulation luminanceModulation1 = new A.LuminanceModulation() { Val = 110000 };
                A.SaturationModulation saturationModulation1 = new A.SaturationModulation() { Val = 105000 };
                A.Tint tint1 = new A.Tint() { Val = 67000 };

                schemeColor2.Append(luminanceModulation1);
                schemeColor2.Append(saturationModulation1);
                schemeColor2.Append(tint1);

                gradientStop1.Append(schemeColor2);

                A.GradientStop gradientStop2 = new A.GradientStop() { Position = 50000 };

                A.SchemeColor schemeColor3 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
                A.LuminanceModulation luminanceModulation2 = new A.LuminanceModulation() { Val = 105000 };
                A.SaturationModulation saturationModulation2 = new A.SaturationModulation() { Val = 103000 };
                A.Tint tint2 = new A.Tint() { Val = 73000 };

                schemeColor3.Append(luminanceModulation2);
                schemeColor3.Append(saturationModulation2);
                schemeColor3.Append(tint2);

                gradientStop2.Append(schemeColor3);

                A.GradientStop gradientStop3 = new A.GradientStop() { Position = 100000 };

                A.SchemeColor schemeColor4 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
                A.LuminanceModulation luminanceModulation3 = new A.LuminanceModulation() { Val = 105000 };
                A.SaturationModulation saturationModulation3 = new A.SaturationModulation() { Val = 109000 };
                A.Tint tint3 = new A.Tint() { Val = 81000 };

                schemeColor4.Append(luminanceModulation3);
                schemeColor4.Append(saturationModulation3);
                schemeColor4.Append(tint3);

                gradientStop3.Append(schemeColor4);

                gradientStopList1.Append(gradientStop1);
                gradientStopList1.Append(gradientStop2);
                gradientStopList1.Append(gradientStop3);
                A.LinearGradientFill linearGradientFill1 = new A.LinearGradientFill() { Angle = 5400000, Scaled = false };

                gradientFill1.Append(gradientStopList1);
                gradientFill1.Append(linearGradientFill1);

                A.GradientFill gradientFill2 = new A.GradientFill() { RotateWithShape = true };

                A.GradientStopList gradientStopList2 = new A.GradientStopList();

                A.GradientStop gradientStop4 = new A.GradientStop() { Position = 0 };

                A.SchemeColor schemeColor5 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
                A.SaturationModulation saturationModulation4 = new A.SaturationModulation() { Val = 103000 };
                A.LuminanceModulation luminanceModulation4 = new A.LuminanceModulation() { Val = 102000 };
                A.Tint tint4 = new A.Tint() { Val = 94000 };

                schemeColor5.Append(saturationModulation4);
                schemeColor5.Append(luminanceModulation4);
                schemeColor5.Append(tint4);

                gradientStop4.Append(schemeColor5);

                A.GradientStop gradientStop5 = new A.GradientStop() { Position = 50000 };

                A.SchemeColor schemeColor6 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
                A.SaturationModulation saturationModulation5 = new A.SaturationModulation() { Val = 110000 };
                A.LuminanceModulation luminanceModulation5 = new A.LuminanceModulation() { Val = 100000 };
                A.Shade shade1 = new A.Shade() { Val = 100000 };

                schemeColor6.Append(saturationModulation5);
                schemeColor6.Append(luminanceModulation5);
                schemeColor6.Append(shade1);

                gradientStop5.Append(schemeColor6);

                A.GradientStop gradientStop6 = new A.GradientStop() { Position = 100000 };

                A.SchemeColor schemeColor7 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
                A.LuminanceModulation luminanceModulation6 = new A.LuminanceModulation() { Val = 99000 };
                A.SaturationModulation saturationModulation6 = new A.SaturationModulation() { Val = 120000 };
                A.Shade shade2 = new A.Shade() { Val = 78000 };

                schemeColor7.Append(luminanceModulation6);
                schemeColor7.Append(saturationModulation6);
                schemeColor7.Append(shade2);

                gradientStop6.Append(schemeColor7);

                gradientStopList2.Append(gradientStop4);
                gradientStopList2.Append(gradientStop5);
                gradientStopList2.Append(gradientStop6);
                A.LinearGradientFill linearGradientFill2 = new A.LinearGradientFill() { Angle = 5400000, Scaled = false };

                gradientFill2.Append(gradientStopList2);
                gradientFill2.Append(linearGradientFill2);

                fillStyleList1.Append(solidFill1);
                fillStyleList1.Append(gradientFill1);
                fillStyleList1.Append(gradientFill2);

                A.LineStyleList lineStyleList1 = new A.LineStyleList();

                A.Outline outline1 = new A.Outline() { Width = 6350, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

                A.SolidFill solidFill2 = new A.SolidFill();
                A.SchemeColor schemeColor8 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };

                solidFill2.Append(schemeColor8);
                A.PresetDash presetDash1 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };
                A.Miter miter1 = new A.Miter() { Limit = 800000 };

                outline1.Append(solidFill2);
                outline1.Append(presetDash1);
                outline1.Append(miter1);

                A.Outline outline2 = new A.Outline() { Width = 12700, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

                A.SolidFill solidFill3 = new A.SolidFill();
                A.SchemeColor schemeColor9 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };

                solidFill3.Append(schemeColor9);
                A.PresetDash presetDash2 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };
                A.Miter miter2 = new A.Miter() { Limit = 800000 };

                outline2.Append(solidFill3);
                outline2.Append(presetDash2);
                outline2.Append(miter2);

                A.Outline outline3 = new A.Outline() { Width = 19050, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

                A.SolidFill solidFill4 = new A.SolidFill();
                A.SchemeColor schemeColor10 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };

                solidFill4.Append(schemeColor10);
                A.PresetDash presetDash3 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };
                A.Miter miter3 = new A.Miter() { Limit = 800000 };

                outline3.Append(solidFill4);
                outline3.Append(presetDash3);
                outline3.Append(miter3);

                lineStyleList1.Append(outline1);
                lineStyleList1.Append(outline2);
                lineStyleList1.Append(outline3);

                A.EffectStyleList effectStyleList1 = new A.EffectStyleList();

                A.EffectStyle effectStyle1 = new A.EffectStyle();
                A.EffectList effectList1 = new A.EffectList();

                effectStyle1.Append(effectList1);

                A.EffectStyle effectStyle2 = new A.EffectStyle();
                A.EffectList effectList2 = new A.EffectList();

                effectStyle2.Append(effectList2);

                A.EffectStyle effectStyle3 = new A.EffectStyle();

                A.EffectList effectList3 = new A.EffectList();

                A.OuterShadow outerShadow1 = new A.OuterShadow() { BlurRadius = 57150L, Distance = 19050L, Direction = 5400000, Alignment = A.RectangleAlignmentValues.Center, RotateWithShape = false };

                A.RgbColorModelHex rgbColorModelHex11 = new A.RgbColorModelHex() { Val = "000000" };
                A.Alpha alpha1 = new A.Alpha() { Val = 63000 };

                rgbColorModelHex11.Append(alpha1);

                outerShadow1.Append(rgbColorModelHex11);

                effectList3.Append(outerShadow1);

                effectStyle3.Append(effectList3);

                effectStyleList1.Append(effectStyle1);
                effectStyleList1.Append(effectStyle2);
                effectStyleList1.Append(effectStyle3);

                A.BackgroundFillStyleList backgroundFillStyleList1 = new A.BackgroundFillStyleList();

                A.SolidFill solidFill5 = new A.SolidFill();
                A.SchemeColor schemeColor11 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };

                solidFill5.Append(schemeColor11);

                A.SolidFill solidFill6 = new A.SolidFill();

                A.SchemeColor schemeColor12 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
                A.Tint tint5 = new A.Tint() { Val = 95000 };
                A.SaturationModulation saturationModulation7 = new A.SaturationModulation() { Val = 170000 };

                schemeColor12.Append(tint5);
                schemeColor12.Append(saturationModulation7);

                solidFill6.Append(schemeColor12);

                A.GradientFill gradientFill3 = new A.GradientFill() { RotateWithShape = true };

                A.GradientStopList gradientStopList3 = new A.GradientStopList();

                A.GradientStop gradientStop7 = new A.GradientStop() { Position = 0 };

                A.SchemeColor schemeColor13 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
                A.Tint tint6 = new A.Tint() { Val = 93000 };
                A.SaturationModulation saturationModulation8 = new A.SaturationModulation() { Val = 150000 };
                A.Shade shade3 = new A.Shade() { Val = 98000 };
                A.LuminanceModulation luminanceModulation7 = new A.LuminanceModulation() { Val = 102000 };

                schemeColor13.Append(tint6);
                schemeColor13.Append(saturationModulation8);
                schemeColor13.Append(shade3);
                schemeColor13.Append(luminanceModulation7);

                gradientStop7.Append(schemeColor13);

                A.GradientStop gradientStop8 = new A.GradientStop() { Position = 50000 };

                A.SchemeColor schemeColor14 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
                A.Tint tint7 = new A.Tint() { Val = 98000 };
                A.SaturationModulation saturationModulation9 = new A.SaturationModulation() { Val = 130000 };
                A.Shade shade4 = new A.Shade() { Val = 90000 };
                A.LuminanceModulation luminanceModulation8 = new A.LuminanceModulation() { Val = 103000 };

                schemeColor14.Append(tint7);
                schemeColor14.Append(saturationModulation9);
                schemeColor14.Append(shade4);
                schemeColor14.Append(luminanceModulation8);

                gradientStop8.Append(schemeColor14);

                A.GradientStop gradientStop9 = new A.GradientStop() { Position = 100000 };

                A.SchemeColor schemeColor15 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
                A.Shade shade5 = new A.Shade() { Val = 63000 };
                A.SaturationModulation saturationModulation10 = new A.SaturationModulation() { Val = 120000 };

                schemeColor15.Append(shade5);
                schemeColor15.Append(saturationModulation10);

                gradientStop9.Append(schemeColor15);

                gradientStopList3.Append(gradientStop7);
                gradientStopList3.Append(gradientStop8);
                gradientStopList3.Append(gradientStop9);
                A.LinearGradientFill linearGradientFill3 = new A.LinearGradientFill() { Angle = 5400000, Scaled = false };

                gradientFill3.Append(gradientStopList3);
                gradientFill3.Append(linearGradientFill3);

                backgroundFillStyleList1.Append(solidFill5);
                backgroundFillStyleList1.Append(solidFill6);
                backgroundFillStyleList1.Append(gradientFill3);

                formatScheme1.Append(fillStyleList1);
                formatScheme1.Append(lineStyleList1);
                formatScheme1.Append(effectStyleList1);
                formatScheme1.Append(backgroundFillStyleList1);

                themeElements1.Append(colorScheme1);
                themeElements1.Append(fontScheme19);
                themeElements1.Append(formatScheme1);
                A.ObjectDefaults objectDefaults1 = new A.ObjectDefaults();
                A.ExtraColorSchemeList extraColorSchemeList1 = new A.ExtraColorSchemeList();

                A.ExtensionList extensionList1 = new A.ExtensionList();

                A.Extension extension1 = new A.Extension() { Uri = "{05A4C25C-085E-4340-85A3-A5531E510DB2}" };

                OpenXmlUnknownElement openXmlUnknownElement3 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<thm15:themeFamily xmlns:thm15=\"http://schemas.microsoft.com/office/thememl/2012/main\" name=\"Office Theme\" id=\"{62F939B6-93AF-4DB8-9C6B-D6C7DFDC589F}\" vid=\"{4A3C46E8-61CC-4603-A589-7422A47A8E4A}\" />");

                extension1.Append(openXmlUnknownElement3);

                extensionList1.Append(extension1);

                theme1.Append(themeElements1);
                theme1.Append(objectDefaults1);
                theme1.Append(extraColorSchemeList1);
                theme1.Append(extensionList1);

                themePart1.Theme = theme1;
            }

            private void SetPackageProperties(OpenXmlPackage document)
            {
                document.PackageProperties.Creator = "Безвершук Дмитро Олександрович";
                document.PackageProperties.Created = System.Xml.XmlConvert.ToDateTime("2017-04-24T06:58:28Z", System.Xml.XmlDateTimeSerializationMode.RoundtripKind);
                document.PackageProperties.Modified = System.Xml.XmlConvert.ToDateTime("2018-01-18T18:26:20Z", System.Xml.XmlDateTimeSerializationMode.RoundtripKind);
                document.PackageProperties.LastModifiedBy = "Безвершук Дмитро Олександрович";
            }


    }
}