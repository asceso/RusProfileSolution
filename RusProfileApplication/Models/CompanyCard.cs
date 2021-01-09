using System.ComponentModel;

namespace RusProfileApplication.Models
{
    public class CompanyCard
    {
        [DisplayName("№ п/п")]
        public int ID { get; set; }

        [DisplayName("Краткое наименование компании")]
        public string ShortName { get; set; }

        [DisplayName("Полное наименование компании")]
        public string FullName { get; set; }

        [DisplayName("Телефоны")]
        public string Phones { get; set; }

        [DisplayName("Почта")]
        public string Mails { get; set; }

        [DisplayName("Сайт")]
        public string Sites { get; set; }

        [DisplayName("ИНН/КПП")]
        public string INN { get; set; }

        [DisplayName("Основной вид деятельности")]
        public string PrimaryOccupation { get; set; }

        [DisplayName("Статус организации")]
        public string OrganizationStatus { get; set; }
    }
}
