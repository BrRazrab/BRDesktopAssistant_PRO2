using System;
using System.Globalization;
using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace BRDesktopAssistant.Services
{
    public static class DocService
    {
        public static string CreateInvoice(string dir, string supplier, string customer, string item, int qty, decimal price)
        {
            var name = $"Счет_{DateTime.Now:yyyyMMdd_HHmmss}.docx";
            var path = Path.Combine(dir, name);
            using var doc = WordprocessingDocument.Create(path, WordprocessingDocumentType.Document);
            var mainPart = doc.AddMainDocumentPart();
            mainPart.Document = new Document(new Body(
                Title("СЧЕТ НА ОПЛАТУ"),
                P($"Дата: {DateTime.Now:dd.MM.yyyy}"),
                P($"Поставщик: {supplier}"),
                P($"Покупатель: {customer}"),
                P($"Товар: {item}"),
                P($"Количество: {qty}"),
                P($"Цена за единицу: {price.ToString("N2", new CultureInfo("ru-RU"))} ₽"),
                P($"Сумма: {(qty*price).ToString("N2", new CultureInfo("ru-RU"))} ₽")
            ));
            mainPart.Document.Save();
            return path;
        }

        public static string CreateWaybill(string dir, string supplier, string customer, string item, int qty, decimal price)
        {
            var name = $"Накладная_{DateTime.Now:yyyyMMdd_HHmmss}.docx";
            var path = Path.Combine(dir, name);
            using var doc = WordprocessingDocument.Create(path, WordprocessingDocumentType.Document);
            var mainPart = doc.AddMainDocumentPart();
            mainPart.Document = new Document(new Body(
                Title("ТОВАРНАЯ НАКЛАДНАЯ"),
                P($"Дата: {DateTime.Now:dd.MM.yyyy}"),
                P($"Поставщик: {supplier}"),
                P($"Получатель: {customer}"),
                P($"Товар: {item} — {qty} шт × {price.ToString("N2", new CultureInfo("ru-RU"))} ₽"),
                P($"Итого: {(qty*price).ToString("N2", new CultureInfo("ru-RU"))} ₽")
            ));
            mainPart.Document.Save();
            return path;
        }

        public static string CreateAct(string dir, string supplierFio, string customerFio, string serviceName, decimal sum)
        {
            var name = $"Акт_{DateTime.Now:yyyyMMdd_HHmmss}.docx";
            var path = Path.Combine(dir, name);
            using var doc = WordprocessingDocument.Create(path, WordprocessingDocumentType.Document);
            var mainPart = doc.AddMainDocumentPart();
            mainPart.Document = new Document(new Body(
                Title("АКТ ВЫПОЛНЕННЫХ РАБОТ"),
                P($"Дата: {DateTime.Now:dd.MM.yyyy}"),
                P($"Исполнитель: {supplierFio}"),
                P($"Заказчик: {customerFio}"),
                P($"Услуга: {serviceName}"),
                P($"Сумма к оплате: {sum.ToString("N2", new CultureInfo("ru-RU"))} ₽")
            ));
            mainPart.Document.Save();
            return path;
        }

        private static Paragraph Title(string text)
        {
            return new Paragraph(new Run(new Text(text))) {
                ParagraphProperties = new ParagraphProperties(new Justification() { Val = JustificationValues.Center })
            };
        }
        private static Paragraph P(string text) => new Paragraph(new Run(new Text(text)));
    }
}
