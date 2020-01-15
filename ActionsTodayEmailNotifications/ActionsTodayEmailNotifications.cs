using System;
using System.Configuration;
using System.Linq;
using System.Collections.Generic;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Host;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk;
using Microsoft.Crm.Sdk.Messages;
using System.Threading.Tasks;

namespace ActionsTodayEmailNotifications
{
    public static class ActionsTodayEmailNotifications
    {
        public static class ConfigurationApp
        {
            public static string crmUser = ConfigurationManager.AppSettings["CrmLogin"];
            public static string crmPassword = ConfigurationManager.AppSettings["CrmPassword"];
            public static string crmUrl = ConfigurationManager.AppSettings["CrmUrl"];
            public static AuthService crm = new AuthService(crmUrl);
            public static IOrganizationService svc = crm.Connect(crmUser, crmPassword);
            public static List<string> entities = new List<string>() { "task", "appointment", "phonecall" };
            public const string Subject = "Subject";
            public static Guid Sender = new Guid("3C5899A9-A855-47DA-B852-88807460807C");
        }

        [FunctionName("ActionsTodayEmailNotifications")]
        public static void Run([TimerTrigger("0 0 10,17 * * *")]TimerInfo myTimer, TraceWriter log)
        {
            IEnumerable<Guid> users = GetActiveUsers();

            foreach (Guid user in users)
            {
                Email email = new Email(ConfigurationApp.Subject, user);

                Parallel.ForEach(ConfigurationApp.entities, (item) =>
                {
                    List<Entity> entityCollection = GetRecords(item, user);
                    email.Add(entityCollection);
                });

                email.Send();
            }
        }

        /// <summary>
        /// Получить всех активных пользователей
        /// </summary>
        /// <returns>Возвращает коллекцию Id пользователей</returns>
        private static IEnumerable<Guid> GetActiveUsers()
        {
            string fetchQuery = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
              <entity name='systemuser'>
                <attribute name='systemuserid' />
                <order attribute='fullname' descending='false' />
                <filter type='and'>
                  <condition attribute='isdisabled' operator='eq' value='0' />
                  <condition attribute='mtr_isbansendemail' operator='eq' value='0' />
                </filter>
              </entity>
            </fetch>";

            return ConfigurationApp.svc.RetrieveMultiple(new FetchExpression(fetchQuery)).Entities.ToList().Select(entity => entity.Id);
        }

        /// <summary>
        /// Забрать все записи по названию сущности и пользователю
        /// </summary>
        /// <param name="logicalName">Название сущности</param>
        /// <param name="user">Id пользователя</param>
        /// <returns>Возвращает коллекцию объектов</returns>
        private static List<Entity> GetRecords(string logicalName, Guid user)
        {
            string fetchQuery = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
              <entity name='{logicalName}'>
                <order attribute='subject' descending='false' />
                <filter type='and'>
                  <condition attribute='ownerid' operator='eq' uitype='systemuser' value='{user}' />
                  <condition attribute='scheduledstart' operator='today' />
                  <condition attribute='statecode' operator='eq' value='0' />
                </filter>
              </entity>
            </fetch>";

            return ConfigurationApp.svc.RetrieveMultiple(new FetchExpression(fetchQuery)).Entities.ToList();
        }

        /// <summary>
        /// Класс для работы с email
        /// </summary>
        internal class Email
        {
            List<string> Messages { get; set; }
            string Subject { get; set; }
            internal string Body { get => string.Join("</br>", Messages); }
            internal Guid Recipient { get; set; }
            internal Guid Sender => ConfigurationApp.Sender;

            public Email(string subject, Guid recipient)
            {
                Subject = subject;
                Recipient = recipient;
                Messages = new List<string>();
            }

            internal void Add(List<Entity> entities)
            {
                if (entities.Any())
                {
                    foreach (Entity entity in entities)
                    {
                        Messages.Add($"{entity.GetAttributeValue<string>("subject")}, Тип действия {GetTypeActity(entity.LogicalName)}, Дата начала {entity.GetAttributeValue<DateTime>("scheduledstart").AddHours(3)}. Ссылка на действие <a href=\"{GenerateUrl(entity.LogicalName, entity.Id)}\">{GenerateUrl(entity.LogicalName, entity.Id)}</a>");
                    }
                }
            }

            internal void Send()
            {
                if (Messages.Any())
                {
                    Guid emailId = CreateEmail();

                    SendEmailRequest sendEmailRequest = new SendEmailRequest
                    {
                        EmailId = emailId,
                        TrackingToken = "",
                        IssueSend = true
                    };

                    ConfigurationApp.svc.Execute(sendEmailRequest);
                }
            }

            private Guid CreateEmail()
            {
                Entity email = ConfigurationEmail();
                Guid emailId = ConfigurationApp.svc.Create(email);
                return emailId;
            }

            private Entity ConfigurationEmail()
            {
                Entity entityFrom = new Entity("activityparty");
                Entity entityTo = new Entity("activityparty");
                entityFrom["partyid"] = new EntityReference("systemuser", Sender);
                entityTo["partyid"] = new EntityReference("systemuser", Recipient);
                Entity email = new Entity("email");
                email["from"] = new Entity[] { entityFrom };
                email["to"] = new Entity[] { entityTo };
                email["subject"] = Subject;
                email["description"] = Body;
                email["directioncode"] = true;
                return email;
            }

            string GenerateUrl(string logicalName, Guid record) => ConfigurationApp.crmUrl.Replace("XRMServices/2011/Organization.svc", $"main.aspx?etn={logicalName}&id={record}&pagetype=entityrecord");

            string GetTypeActity(string logicalName)
            {
                switch (logicalName)
                {
                    case "phonecall":
                        return "Звонок";
                    case "task":
                        return "Задача";
                    case "appointment":
                        return "Встреча";
                    default:
                        throw new ArgumentException("message", nameof(logicalName));
                }
            }
        }
    }
}
