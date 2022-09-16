using System.Diagnostics;
using GemBox.Presentation;
using System.Drawing;
using GemBox.Spreadsheet.Charts;
using GemBox.Spreadsheet;
using ColorName = GemBox.Presentation.ColorName;
using Color = GemBox.Presentation.Color;
using System.Text;
using LengthUnit = GemBox.Presentation.LengthUnit;
using Org.BouncyCastle.Crypto;
using Org.BouncyCastle.Asn1.X500;
using System.Threading;
using RestSharp;
using RestSharp.Authenticators;
using static System.Net.WebRequestMethods;
using Org.BouncyCastle.Utilities;
using System.Xml;
using Newtonsoft.Json;
using Formatting = Newtonsoft.Json.Formatting;
using static System.Net.Mime.MediaTypeNames;
using Org.BouncyCastle.Asn1.Ocsp;
using GemBox.Presentation;
using ShapeCrawler;
using GemBox.Spreadsheet.Drawing;
using DocumentFormat.OpenXml.Spreadsheet;
using System;

ComponentInfo.SetLicense("FREE-LIMITED-KEY");
SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");

int id = 0;

string code = "03cbf5ab-1caa-4a44-94b7-2505263aa828";

RestClient restClient = new("https://hangwire.kimola.com/v1/");

string json = "{\"content\":null,\"models\":{},\"pieces\":{},\"topics\":[],\"categories\":[],\"from\":0,\"size\":0}";

RestRequest restRequest = new RestRequest("reports/{code}/rows/statistics", Method.Post);
restRequest.AddParameter("application/json", json, ParameterType.RequestBody);
restRequest.AddUrlSegment("code", code);

RestResponse<dynamic> statisticsJson = restClient.Execute<dynamic>(restRequest);

restRequest = new RestRequest("reports/{code}", Method.Get);
restRequest.AddUrlSegment("code", code);

RestResponse<dynamic> reportJson = restClient.Execute<dynamic>(restRequest);

dynamic report = JsonConvert.DeserializeObject(reportJson.Content);
dynamic statistics = JsonConvert.DeserializeObject(statisticsJson.Content);

List<PresentationDocument> presentations = new();
//yeni bir sunum
var presentation = new PresentationDocument();
presentations.Add(presentation);

// Add new PowerPoint presentation slide.
var slide = presentation.Slides.AddNew(SlideLayoutType.Blank);

var quarterCircle = slide.Content.AddShape(
      ShapeGeometryType.Donut , -3.3, -3.3, 6.84, 6.84, LengthUnit.Centimeter);

// Get shape outline format.
var lineFormat = quarterCircle.Format;

// Get shape fill format.
var fillFormat = lineFormat.Fill;

// Set shape fill format as solid fill.
fillFormat.SetSolid(Color.FromRgb(47,109,255));

var timelineTrend = slide.Content.AddTextBox(
 ShapeGeometryType.RoundedRectangle, 2.53, 2.4, 8, 2, LengthUnit.Centimeter);
var timelineTrendBox = timelineTrend.AddParagraph().AddRun("Timeline Trend");
timelineTrendBox.Format.Size = 32;
timelineTrendBox.Format.Bold = true;

var underTimelineTrend = slide.Content.AddTextBox(
 ShapeGeometryType.RoundedRectangle, 2.04, 4.42, 11.58, 2, LengthUnit.Centimeter);
var boxUnderTimelineTrend = underTimelineTrend.AddParagraph().AddRun("See how the volume of the conversations evolved among time.");
boxUnderTimelineTrend.Format.Size = 14;

//--------sağ üst sarı olan
var verticalLine1 = slide.Content.AddShape(ShapeGeometryType.Rectangle,20.25, 1.93, 0.47, 2.83, GemBox.Presentation.LengthUnit.Centimeter);
var formatLine1 = verticalLine1.Format;
var fillFormatLine1 = formatLine1.Fill;
verticalLine1.Format.Fill.SetSolid(Color.FromRgb(255, 214, 28));
verticalLine1.Format.Outline.Fill.SetNone();
//---------sağ üst sarı olan

var popularTerms = slide.Content.AddTextBox(
 ShapeGeometryType.RoundedRectangle, 20.99, 1.71, 7.55, 1.63, LengthUnit.Centimeter);
var boxPopularTerms = popularTerms.AddParagraph().AddRun("Popular Terms");
boxPopularTerms.Format.Size = 32;
boxPopularTerms.Format.Bold = true;

var underPopularTerms = slide.Content.AddTextBox(
 ShapeGeometryType.RoundedRectangle, 20.93, 3.31, 9.48, 1.62, LengthUnit.Centimeter);
var boxUnderPopularTerms = underPopularTerms.AddParagraph().AddRun("Here are the popular terms that are mentioned.");
boxUnderPopularTerms.Format.Size = 14;

//--------sağ alt kırmızı olan
var verticalLine2 = slide.Content.AddShape(ShapeGeometryType.Rectangle,20.23, 9.09, 0.47, 2.83, GemBox.Presentation.LengthUnit.Centimeter);
var formatLine2 = verticalLine2.Format;
var fillFormatLine2 = formatLine2.Fill;
verticalLine2.Format.Fill.SetSolid(Color.FromRgb(255, 87, 86));
verticalLine2.Format.Outline.Fill.SetNone();
//---------sağ alt kırmızı olan

var popularCategories = slide.Content.AddTextBox(
 ShapeGeometryType.RoundedRectangle, 20.99, 8.72, 9.78, 1.63, LengthUnit.Centimeter);
var boxPopularCategories = popularCategories.AddParagraph().AddRun("Popular Categories");
boxPopularCategories.Format.Size = 32;
boxPopularCategories.Format.Bold = true;

var underPopularCategories = slide.Content.AddTextBox(
 ShapeGeometryType.RoundedRectangle, 21.11, 10.17, 9.3, 1.45, LengthUnit.Centimeter);
var boxUnderPopularCategories = underPopularCategories.AddParagraph().AddRun("Here are the popular categories classified with our NLP technology. ");
boxUnderPopularCategories.Format.Size = 14;

if (statistics.dates.Count != 0)
{
  var chart = slide.Content.AddChart(GemBox.Presentation.ChartType.Column,1.35, 6.23, 18.57, 11.4, GemBox.Presentation.LengthUnit.Centimeter);
  // Get underlying Excel chart.
  ExcelChart excelChart = (ExcelChart)chart.ExcelChart;
  ExcelWorksheet worksheet = excelChart.Worksheet;

  int i = 2;
  foreach (var date in statistics.dates)
  {
    string dateTime = date.date;
    int count = date.count;
    dateTime = dateTime.Split(" ")[0];
    worksheet.Cells[$"A{i}"].Value = dateTime;
    worksheet.Cells[$"B{i}"].Value = count;
    i++;
  }
  excelChart.SelectData(worksheet.Cells.GetSubrange($"A{1}:B{i}"), true);
}
else
{
  var textBoxEmpty = slide.Content.AddTextBox(
  ShapeGeometryType.RoundedRectangle, 1.51, 8, 20, 3, LengthUnit.Centimeter);
  var emptyBox = textBoxEmpty.AddParagraph().AddRun("There is no chart here !.");
  emptyBox.Format.Size = 28;
}

RestRequest restRequestSections = new("reports/{code}/sections", Method.Get);
restRequestSections.AddUrlSegment("code", code);
RestResponse<dynamic> restResponseSections = restClient.Execute<dynamic>(restRequestSections);

dynamic categoriesArray = JsonConvert.DeserializeObject(restResponseSections.Content);

var pieChart = slide.Content.AddChart(GemBox.Presentation.ChartType.Pie,21.1, 11.92, 9.74, 6.32, GemBox.Presentation.LengthUnit.Centimeter);

ExcelChart excelPieChart = (ExcelChart)pieChart.ExcelChart;
ExcelWorksheet worksheet1 = excelPieChart.Worksheet;

int s = 2;
int circleCount = 1;
foreach (dynamic item in categoriesArray)
{
  if (item.name == "Popular Categories")
  {
    foreach (dynamic itemSubCategory in item.items)
    {
      if (circleCount == 5)
        break;
      string name1 = itemSubCategory.name;
      int count = itemSubCategory.count;
      worksheet1.Cells[$"A{s}"].Value = count;
      worksheet1.Cells[$"B{s}"].Value = name1;
      s++;
      circleCount++; 
    }
  }
}

excelPieChart.SelectData(worksheet1.Cells.GetSubrange($"A1:B{s}"), true, true);
excelPieChart.DataLabels.LabelContainsPercentage = true;
excelPieChart.DataLabels.LabelContainsValue = false;
excelPieChart.DataLabels.LabelPosition = DataLabelPosition.Center;

var underPopularTerms2 = slide.Content.AddTextBox(ShapeGeometryType.RoundedRectangle, 20.95, 4.75, 9.89, 3.43, LengthUnit.Centimeter);

string boxString = "";
foreach (var topic in statistics.topics)
{
  string topicName = topic.name;
  string topicCount = topic.count;
  string topicPercentage = topic.percentage;
  boxString += topicName + ", ";
}
var boxUnderPopularTerms2 = underPopularTerms2.AddParagraph().AddRun(boxString);
boxUnderPopularTerms2.Format.Size = 13;

foreach (var model in statistics.models)
{
  string modelName = model.name;
  //----------TITLE----------
  slide = presentation.Slides.AddNew(SlideLayoutType.Blank);

  string title2 = modelName;
  var textBox2 = slide.Content.AddTextBox(ShapeGeometryType.Rectangle,2, 2, 30, 3, GemBox.Presentation.LengthUnit.Centimeter);
  var titleTextBox2 = textBox2.AddParagraph().AddRun(title2);
  titleTextBox2.Format.Size = 28;

  //graph positions
  double graphPositionX = 2;
  double graphPositionY = 10;
  double graphWidthX = 10;

  //words positions
  double wordPositionX = 2;
  double wordPositionY = 9;

  int count = 1; //sayfada 4 tane grafik oluşturmayı sağlıyor.

  double OrangePercantage;

  foreach (var label in model.statistics)
  {
    string labelName = label.name;
    int labelCount = label.count;
    int labelPercentage = label.percentage;
    if (count == 5)
      break;

    double DlabelPercentage = Convert.ToDouble(labelPercentage); //83
    OrangePercantage = graphWidthX * (DlabelPercentage / 100);
    //graph start
    var shape = slide.Content.AddShape(ShapeGeometryType.Rectangle,graphPositionX, graphPositionY, graphWidthX, 0.7, GemBox.Presentation.LengthUnit.Centimeter);

    var format1 = shape.Format;
    var fillFormat1 = format1.Fill;

    shape.Format.Fill.SetSolid(Color.FromRgb(242, 242, 242));
    shape.Format.Outline.Fill.SetNone();
    //graph end

    string label1 = labelName;
    //----------TİTLE sentiment adı----------
    var labelTextBox1 = slide.Content.AddTextBox(ShapeGeometryType.Rectangle, wordPositionX, wordPositionY, 3.5, 0.7, GemBox.Presentation.LengthUnit.Centimeter);
    labelTextBox1.Format.WrapText = false;

    var leftLabelTextBox = labelTextBox1.AddParagraph().AddRun(label1);
    leftLabelTextBox.Format.Size = 20;
    //-----------title weak words end

    string label3 = $"{labelPercentage}% ({labelCount.ToString("N0")})";
    //----------title percentage words start----------
    var labelTextBox3 = slide.Content.AddTextBox(ShapeGeometryType.Rectangle, wordPositionX + 10, wordPositionY + 1, 1.3, 0.7, GemBox.Presentation.LengthUnit.Centimeter);
    labelTextBox3.Format.WrapText = false;

    var boyut = labelTextBox3.AddParagraph().AddRun(label3);
    boyut.Format.Size = 12;
    //-----------title percentage words end.

    //graph start
    var shape1 = slide.Content.AddShape(ShapeGeometryType.Rectangle,graphPositionX, graphPositionY, OrangePercantage, 0.7,GemBox.Presentation.LengthUnit.Centimeter);
    var format2 = shape.Format;
    var fillFormat2 = format2.Fill;

    shape1.Format.Fill.SetSolid(Color.FromRgb(255, 192, 0));
    shape1.Format.Outline.Fill.SetNone();
    //graph end

    //graph positions
    graphPositionX = 2;
    graphPositionY = graphPositionY + 2;

    //words positions
    wordPositionX = 2;
    wordPositionY = wordPositionY + 2;
    //--------arada kalan vertical line bölümü başlangıç
    var shapeLine = slide.Content.AddShape(ShapeGeometryType.Rectangle,16.5, 9, 0.01, 9, GemBox.Presentation.LengthUnit.Centimeter);

    var formatLine = shapeLine.Format;
    var fillFormatLine = formatLine.Fill;
    shapeLine.Format.Fill.SetSolid(Color.FromRgb(127, 127, 127));
    shapeLine.Format.Fill.SetNone();
    //---------vertical line bölüm bitiş
    count++;
  }

  if (presentation.Slides.Count == 3)
  {
    presentation = new PresentationDocument();
    presentations.Add(presentation);
  }
}

//      TODO
//sağ grafik alanı için api hazır değil o yüzden boş

double x2Degiskeni = 2; // bu da diğer değişkenlerin yuzdelerine bağlanacak.
int kacGrafikVar = 4; //buna gerek kalmayacak foreach ile dondugumuzde hallolacak


////foreach (var item in collection)
////{
////  ------------sağ label grafik alanı başlangıc
////int sağkacGrafikVar = 4;

////  graph positions
////double rightGraphPositionX = 20;
////  double rightGraphPositionY = 10;

////  words positions
////double rightWordPositionX = 20;
////  double rightWordPositionY = 9;

////  graph start
////  var shapes1 = slide.Content.AddShape(
////      ShapeGeometryType.Rectangle,
////             rightGraphPositionX, rightGraphPositionY, 10, 0.7, GemBox.Presentation.LengthUnit.Centimeter);

////  var formats1_1 = shapes1.Format;
////  var fillFormats1_1 = formats1_1.Fill;
////  shapes1.Format.Fill.SetSolid(Color.FromRgb(127, 127, 127));
////  shapes1.Format.Outline.Fill.SetNone();


////  graph end

////  TODO buradaki x genişliği değişebilir olması lazım ki değere göre değişmiş olsun


////  string labels1_1 = "sentiment adı ";
////  ----------TİTLE sentiment adı----------
////  var labelTextBoxs1_1 = slide.Content.AddTextBox(ShapeGeometryType.Rectangle,
////      rightWordPositionX, rightWordPositionY, 4.7, 0.7, GemBox.Presentation.LengthUnit.Centimeter);

////  var rightLabelTextboxs1_1 = labelTextBoxs1_1.AddParagraph().AddRun(labels1_1);
////  rightLabelTextboxs1_1.Format.Size = 20;

////  -----------title strong words start

////  string labels2_1 = "silik yazı ";
////  ----------title weak words start----------
////  var labelTextBoxs2_1 = slide.Content.AddTextBox(ShapeGeometryType.Rectangle,
////      rightWordPositionX + 4.2, rightWordPositionY + 0.15, 3, 1, GemBox.Presentation.LengthUnit.Centimeter);

////  var rightLabel1 = labelTextBoxs2_1.AddParagraph().AddRun(labels2_1);
////  rightLabel1.Format.Fill.SetSolid(Color.FromName(ColorName.Gray));
////  rightLabel1.Format.Size = 16;

////  -----------title weak words end

////  string labels3_1 = "%43 ";
////  ----------title percentage words start----------
////  var labelTextBoxs3_1 = slide.Content.AddTextBox(ShapeGeometryType.Rectangle,
////      rightWordPositionX + 10, rightWordPositionY + 0.85, 2.5, 1, GemBox.Presentation.LengthUnit.Centimeter);

////  var boyuts1 = labelTextBoxs3_1.AddParagraph().AddRun(labels3_1);
////  boyuts1.Format.Size = 16;

////  -----------title percentage words end

////  graph start

////  var shapes1_4 = slide.Content.AddShape(
////  ShapeGeometryType.Rectangle,
////         rightGraphPositionX, rightGraphPositionY, xDegiskeni, 0.7, GemBox.Presentation.LengthUnit.Centimeter);
////  var formats1_4 = shapes1_4.Format;
////  var fillFormats1_4 = formats1_4.Fill;

////  shapes1_4.Format.Fill.SetSolid(Color.FromRgb(155, 187, 89));
////  shapes1_4.Format.Outline.Fill.SetNone();

////  graph end

////  2.grafik başlangıcı
////  var shapes1_5 = slide.Content.AddShape(
//// ShapeGeometryType.Rectangle,
////        xDegiskeni + rightWordPositionX, rightGraphPositionY, x2Degiskeni, 0.7, GemBox.Presentation.LengthUnit.Centimeter);
////  var formats1_5 = shapes1_5.Format;
////  var fillFormats1_5 = formats1_5.Fill;

////  shapes1_5.Format.Fill.SetSolid(Color.FromRgb(149, 55, 53));
////  shapes1_5.Format.Outline.Fill.SetNone();


////  graph positions
////  rightGraphPositionX = 20;
////  rightGraphPositionY = rightGraphPositionY + 2;

////  words positions
////  rightWordPositionX = 20;
////  rightWordPositionY = rightWordPositionY + 2;

////  2.graph end


////-------- - sağ label grafik alanı bitiş
////}

RestClient restClient1 = new("https://machineheart.kimola.com/v1/");
id = report.id;
RestRequest restRequest1 = new RestRequest($"Members/{id}", Method.Get);
restRequest1.AddUrlSegment("id", id);
RestResponse<dynamic> members = restClient1.Execute<dynamic>(restRequest1);

dynamic member2 = JsonConvert.DeserializeObject(members.Content);

string name = member2.firstName;
string lastName = member2.lastName;
string nameTitle = report.title;

string nowDate = DateTime.Now.ToString("M/d/yyyy");

var presentationKapak = PresentationDocument.Load("Giris.pptx");
var slide0 = presentationKapak.Slides[0];
slide0.Content.Drawings[6].TextContent.Replace("SAMPLE REPORT \nNAME TITLE V23\n", $"{nameTitle}");
slide0.Content.Drawings[7].TextContent.Replace("This report is created by USERNAME on 24.03.2022 with Kimola Cognitive.\n", $"This report is created by {name} {lastName} on {nowDate} with Kimola Cognitive.\n");

presentationKapak.Save("Giris1.pptx");

List<string> files = new List<string>();
for (int n = 0; n < presentations.Count; n++)
{
  presentations[n].Save($"chart-{n}.pptx");
  files.Add($"chart-{n}.pptx");
}

using var finalPre = SCPresentation.Open(@"Giris1.pptx", true);
foreach (var file in files)
{
  using var sourcePre = SCPresentation.Open(file, false);
  var pageNumber = sourcePre.Slides.Count;
  for (int l = 0; l < pageNumber; l++)
  {
    var copyingSlide = sourcePre.Slides[l];
    finalPre.Slides.Add(copyingSlide);
  }
}