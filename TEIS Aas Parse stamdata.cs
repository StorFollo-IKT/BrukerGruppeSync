using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Xml.Serialization;

namespace AsAgressoXMLToCSV
{
    [Guid("6C5E9E79-2B68-41B1-BD6A-05EBEA29C139")]
    [ClassInterface(ClassInterfaceType.AutoDual)]
    [ComVisible(true)]
    public class Program : IProgram
    {
        public static Dictionary<string, string> dic = new Dictionary<string, string>();
        public static Dictionary<string, string> groups = new Dictionary<string, string>();
        public List<string> AllUsers = new List<string>();

        public void Main(string inXML, string outCSV, string inGroupFile)
        {
            ExportInfo exportInfo = new ExportInfo();
            using (StreamReader streamReader = new StreamReader(inXML, Encoding.UTF8))
                exportInfo = (ExportInfo)new XmlSerializer(typeof(ExportInfo)).Deserialize((TextReader)streamReader);
            if (File.Exists(outCSV))
                File.Delete(outCSV);
            File.AppendAllText(outCSV, "Action;givenname;sn;SSN;ResourceID;JobTitle;Department;DepartmentID;PostIdDescription;Manager;OtherPostIDs;Office;Groups;R_Employments;error;comment" + Environment.NewLine, Encoding.GetEncoding("iso-8859-1"));
            using (StreamReader streamReader = new StreamReader(inGroupFile, Encoding.GetEncoding("iso-8859-1")))
            {
                while (!streamReader.EndOfStream)
                {
                    IEnumerable<string> source = (IEnumerable<string>)streamReader.ReadLine().Split(';');
                    if (Program.groups.ContainsKey(source.ElementAt<string>(0)))
                    {
                        Dictionary<string, string> groups = Program.groups;
                        string key = source.ElementAt<string>(0);
                        groups[key] = groups[key] + "|" + string.Join("|", source.Skip<string>(1).ToArray<string>());
                    }
                    else
                        Program.groups.Add(source.ElementAt<string>(0), string.Join("|", source.Skip<string>(1).ToArray<string>()));
                }
            }
            foreach (ExportInfoOrganisation organisation in exportInfo.Organisations)
            {
                if (!Program.dic.ContainsKey(organisation.Id))
                    Program.dic.Add(organisation.Id, organisation.Managers.@string);
            }
            foreach (ExportInfoResource resource in exportInfo.Resources)
            {
                Person tPerson = new Person()
                {
                    FName = resource.FirstName,
                    LName = resource.Surname,
                    SSN = resource.SocialSecurityNumber,
                    ResourceID = resource.ResourceId,
                    L_Employments = new List<PersonEmployments>(),
                    Groups = new List<string>()
                };
                string AllEmploymentCodes = string.Join("", ((IEnumerable<ExportInfoResourceEmployment>)resource.Employments).Select<ExportInfoResourceEmployment, string>((Func<ExportInfoResourceEmployment, string>)(employment => employment.EmploymentType)));
                foreach (ExportInfoResourceEmployment employment in resource.Employments)
                {
                    switch (Program.isActiveEmployment(employment, AllEmploymentCodes))
                    {
                        case Program.EmploymentStatus.InActive:
                            continue;
                        case Program.EmploymentStatus.IgnoreUser:
                            tPerson.Ignored = true;
                            ref string local = ref tPerson.Comment;
                            local = local + "Users was ignored since it had a Employment with a specific postcode : (" + employment.PostCode + ") |";
                            tPerson.L_Employments.Clear();
                            goto label_32;
                        default:
                            foreach (ExportInfoResourceEmploymentRelation relation in employment.Relations)
                            {
                                string str = "";
                                if (employment.MainPosition)
                                    Program.dic.TryGetValue(relation.Value, out str);
                                List<PersonEmployments> lEmployments = tPerson.L_Employments;
                                PersonEmployments personEmployments = new PersonEmployments();
                                personEmployments.Name = relation.Name;
                                personEmployments.Value = relation.Value;
                                personEmployments.Description = relation.Description;
                                personEmployments.manager = str;
                                personEmployments.Percentage = employment.Percentage;
                                personEmployments.postIdDescription = employment.PostIdDescription;
                                personEmployments.ElementType = relation.ElementType;
                                personEmployments.DateFrom = relation.DateFrom.ToShortDateString();
                                DateTime dateTime = relation.DateTo;
                                dateTime = dateTime.AddDays(1.0);
                                personEmployments.DateTo = dateTime.ToShortDateString();
                                personEmployments.MainPos = employment.MainPosition;
                                personEmployments.JobTitle = employment.PostCodeDescription;
                                lEmployments.Add(personEmployments);
                            }
                            continue;
                    }
                }
            label_32:
                Program.HandleEmployments(ref tPerson, true);
                this.AddUserToOutputList(ref tPerson);
            }
            this.PrintAllUsersInOutputList(outCSV);
        }

        private void PrintAllUsersInOutputList(string inCSV)
        {
            this.AllUsers.Sort();
            foreach (string allUser in this.AllUsers)
                File.AppendAllText(inCSV, allUser, Encoding.GetEncoding("iso-8859-1"));
        }

        private static Program.EmploymentStatus isActiveEmployment(
          ExportInfoResourceEmployment tEmployment,
          string AllEmploymentCodes)
        {
            if (tEmployment.PostCode == "999999" && !AllEmploymentCodes.Contains("A") || tEmployment.PostCode == "21400" || tEmployment.PostCode == "40500" || tEmployment.PostCode == "51000" || tEmployment.PostCode == "51100" || tEmployment.PostCode == "51200" || tEmployment.PostCode == "60600" || tEmployment.PostCode == "70900" || tEmployment.PostCode == "80101" || tEmployment.PostCode == "90700")
                return Program.EmploymentStatus.IgnoreUser;
            if (tEmployment.Percentage > 0M)
                return Program.EmploymentStatus.Active;
            switch (tEmployment.EmploymentType.ToUpper())
            {
                case "A":
                    return Program.EmploymentStatus.Active;
                case "B":
                    return Program.EmploymentStatus.Active;
                case "C":
                    return Program.EmploymentStatus.Active;
                case "D":
                    return Program.EmploymentStatus.Active;
                case "E":
                    return Program.EmploymentStatus.Active;
                case "F":
                    return Program.EmploymentStatus.Active;
                case "G":
                    return Program.EmploymentStatus.Active;
                case "H":
                    return Program.EmploymentStatus.InActive;
                case "I":
                    return Program.EmploymentStatus.InActive;
                case "J":
                    return Program.EmploymentStatus.Active;
                case "K":
                    return Program.EmploymentStatus.Active;
                case "L":
                    return Program.EmploymentStatus.Active;
                case "M":
                    return Program.EmploymentStatus.Active;
                case "N":
                    return Program.EmploymentStatus.Active;
                case "O":
                    return Program.EmploymentStatus.Active;
                case "P":
                    return Program.EmploymentStatus.Active;
                case "Q":
                    return Program.EmploymentStatus.Active;
                case "R":
                    return Program.EmploymentStatus.Active;
                case "S":
                    return Program.EmploymentStatus.Active;
                case "T":
                    return Program.EmploymentStatus.Active;
                case "U":
                    return Program.EmploymentStatus.Active;
                case "V":
                    return Program.EmploymentStatus.InActive;
                case "W":
                    return Program.EmploymentStatus.Active;
                case "X":
                    return Program.EmploymentStatus.Active;
                case "Y":
                    return Program.EmploymentStatus.InActive;
                case "Z":
                    return Program.EmploymentStatus.InActive;
                case "Æ":
                    return Program.EmploymentStatus.Active;
                default:
                    throw new InvalidDataException("Employment type: " + tEmployment.EmploymentType + " is not among the defaults, A-Z + Æ");
            }
        }

        private static void HandleEmployments(ref Person tPerson, bool forcemainPos)
        {
            tPerson.Groups.Add(Program.groups["everyone"]);
            if (tPerson.L_Employments.Count > 0)
            {
                tPerson.L_Employments = tPerson.L_Employments.OrderByDescending<PersonEmployments, Decimal>((Func<PersonEmployments, Decimal>)(inO => inO.Percentage)).ToList<PersonEmployments>();
                bool flag = true;
                foreach (PersonEmployments lEmployment in tPerson.L_Employments)
                {
                    if (flag && Convert.ToDateTime(lEmployment.DateTo) > DateTime.Today)
                        flag = false;
                    ref string local1 = ref tPerson.R_Employments;
                    local1 = local1 + lEmployment.Name + "|" + lEmployment.Value + "|" + lEmployment.Description + "|" + lEmployment.manager + "|" + (object)lEmployment.Percentage + "%|" + lEmployment.postIdDescription + "|" + lEmployment.ElementType + "|From:" + lEmployment.DateFrom + "|To:" + lEmployment.DateTo + "&";
                    if (lEmployment.Name == "Ansvar" && Convert.ToDateTime(lEmployment.DateTo) > DateTime.Today && !tPerson.AnsvarSet && (lEmployment.MainPos || !forcemainPos))
                    {
                        tPerson.AnsvarSet = true;
                        tPerson.Department = lEmployment.Description;
                        tPerson.DepartmentID = lEmployment.Value;
                        tPerson.PostIdDescription = lEmployment.postIdDescription;
                        tPerson.JobTitle = lEmployment.JobTitle;
                    }
                    else if (lEmployment.Name == "Arbeidssted")
                    {
                        ref string local2 = ref tPerson.Office;
                        local2 = local2 + lEmployment.Description + "|";
                        ref string local3 = ref tPerson.OtherPostIDs;
                        local3 = local3 + lEmployment.Name + "=" + lEmployment.Value + "=" + (object)lEmployment.Percentage + "%=" + lEmployment.postIdDescription + "=";
                    }
                    else if (!string.IsNullOrEmpty(lEmployment.manager) && (string.IsNullOrEmpty(tPerson.Manager) || !tPerson.Manager.Contains(lEmployment.manager)) && (lEmployment.MainPos || !forcemainPos))
                    {
                        ref string local4 = ref tPerson.Manager;
                        local4 = local4 + lEmployment.manager + "|";
                    }
                    if (Program.groups.ContainsKey(lEmployment.Value))
                    {
                        tPerson.Groups.Add(Program.groups[lEmployment.Value] + "|");
                    }
                    else
                    {
                        ref string local5 = ref tPerson.error;
                        local5 = local5 + "<Missing EmploymentValue in groupConfig: " + lEmployment.Value + "> ";
                    }
                }
                if (!tPerson.AnsvarSet & forcemainPos)
                {
                    tPerson.Office = "";
                    tPerson.OtherPostIDs = "";
                    tPerson.Manager = "";
                    tPerson.Groups.Clear();
                    tPerson.error = "";
                    Program.HandleEmployments(ref tPerson, false);
                }
                else
                {
                    tPerson.Action = flag ? "InActive" : "Active";
                    if (tPerson.Ignored)
                        return;
                    tPerson.Comment += flag ? "allJobsDateToLessThanToday" : "";
                }
            }
            else
            {
                tPerson.Action = "InActive";
                if (tPerson.Ignored)
                    return;
                tPerson.Comment += "No Active employments";
            }
        }

        private void AddUserToOutputList(ref Person tPerson)
        {
            if (tPerson.Action == "InActive")
                tPerson.Groups.Clear();
            this.AllUsers.Add(tPerson.Action + ";" + tPerson.FName + ";" + tPerson.LName + ";" + tPerson.SSN + ";" + tPerson.ResourceID + ";" + tPerson.JobTitle + ";" + tPerson.Department + ";" + tPerson.DepartmentID + ";" + tPerson.PostIdDescription + ";" + tPerson.Manager + ";" + tPerson.OtherPostIDs + ";" + tPerson.Office + ";" + string.Join("|", tPerson.Groups.ToArray()) + ";" + tPerson.R_Employments + ";" + ";" + tPerson.Comment + ";" + Environment.NewLine);
        }

        private static void Main() => new Program().Main("C:\\Users\\knut.haug\\OneDrive - StorFollo\\Dokumenter\\stamdata\\Ås\\Stamdata3_FSI_AL.xml", "C:\\Users\\knut.haug\\OneDrive - StorFollo\\Dokumenter\\stamdata\\Ås\\Stamdata3_FSI_AL.csv", "C:\\Users\\knut.haug\\OneDrive - StorFollo\\Dokumenter\\stamdata\\Ås\\Config.txt");

        private enum EmploymentStatus
        {
            Active,
            InActive,
            IgnoreUser,
        }
    }

    [Guid("97897D58-0642-40BE-92B6-27AA8550FE1A")]
    [ComVisible(true)]
    internal interface IProgram
    {
        void Main(string inXML, string outCSV, string inGroupFile);
    }

    public class PersonEmployments
    {
        public string Name;
        public string Value;
        public string Description;
        public string manager;
        public Decimal Percentage;
        public string postIdDescription;
        public string ElementType;
        public string DateFrom;
        public string DateTo;
        public bool MainPos;
        private string _JobTitle;

        public string JobTitle
        {
            get => this._JobTitle;
            set => this._JobTitle = value;
        }
    }

    public struct Person
    {
        public string Action;
        public string FName;
        public string LName;
        public string ResourceID;
        public string SSN;
        public List<PersonEmployments> L_Employments;
        public string R_Employments;
        public string JobTitle;
        public List<string> Groups;
        public string Department;
        public string PostIdDescription;
        public string OtherPostIDs;
        public string Manager;
        public bool AnsvarSet;
        public string error;
        public string DepartmentID;
        public string Office;
        public string Comment;
        public bool Ignored;
    }

    [XmlType(AnonymousType = true)]
    [XmlRoot(IsNullable = false, Namespace = "")]
    public class ExportInfo
    {
        [XmlArrayItem("Resource", IsNullable = false)]
        public ExportInfoResource[] Resources { get; set; }

        public string NumberOfResources { get; set; }

        public DateTime ExportedDate { get; set; }

        [XmlArrayItem("Organisation", IsNullable = false)]
        public ExportInfoOrganisation[] Organisations { get; set; }
    }

    [XmlType(AnonymousType = true)]
    public class ExportInfoOrganisation
    {
        public string CompanyCode { get; set; }

        public string Name { get; set; }

        public string Id { get; set; }

        public string ParentId { get; set; }

        public ExportInfoOrganisationManagers Managers { get; set; }

        public DateTime DateFrom { get; set; }

        public DateTime DateTo { get; set; }

        public DateTime LastUpdate { get; set; }

        public string Status { get; set; }
    }

    [XmlType(AnonymousType = true)]
    public class ExportInfoOrganisationManagers
    {
        public string @string { get; set; }
    }

    [XmlType(AnonymousType = true)]
    public class ExportInfoResource
    {
        [XmlArrayItem("Address", IsNullable = false)]
        public ExportInfoResourceAddress[] Addresses { get; set; }

        [XmlArrayItem("Employment", IsNullable = false)]
        public ExportInfoResourceEmployment[] Employments { get; set; }

        [XmlArrayItem("Relation", IsNullable = false)]
        public ExportInfoResourceRelation[] Relations { get; set; }

        public object Rates { get; set; }

        public string CompanyCode { get; set; }

        public string ResourceId { get; set; }

        [XmlElement(DataType = "date")]
        public DateTime DateFrom { get; set; }

        [XmlElement(DataType = "date")]
        public DateTime DateTo { get; set; }

        public string FirstName { get; set; }

        public string Surname { get; set; }

        public string Name { get; set; }

        public string ShortName { get; set; }

        public string ResourceType { get; set; }

        [XmlElement(DataType = "date")]
        public DateTime Birthdate { get; set; }

        public string SocialSecurityNumber { get; set; }

        public string Sex { get; set; }

        public string Municipal { get; set; }

        public DateTime LastUpdate { get; set; }

        public string Status { get; set; }

        public object Dim1 { get; set; }

        public object Dim2 { get; set; }

        public object Dim3 { get; set; }

        public object Dim4 { get; set; }

        public string OvertimeYtd { get; set; }

        public DateTime TotalLastUpdate { get; set; }
    }

    [XmlType(AnonymousType = true)]
    public class ExportInfoResourceAddress
    {
        public string AddressId { get; set; }

        public string Type { get; set; }

        public string Contact { get; set; }

        public object Position { get; set; }

        public string Street { get; set; }

        public string Place { get; set; }

        public string Province { get; set; }

        public string CountryCode { get; set; }

        public string ZipCode { get; set; }

        public string Telephone { get; set; }

        public object Telefax { get; set; }

        public string Telex { get; set; }

        public string Mobile { get; set; }

        public object Pager { get; set; }

        public string Home { get; set; }

        public object Assistant { get; set; }

        [XmlArrayItem(IsNullable = false)]
        public string[] EMailList { get; set; }

        public ExportInfoResourceAddressEMailCopyList EMailCopyList { get; set; }

        public DateTime LastUpdate { get; set; }
    }

    [XmlType(AnonymousType = true)]
    public class ExportInfoResourceAddressEMailCopyList
    {
        public string @string { get; set; }
    }

    [XmlType(AnonymousType = true)]
    public class ExportInfoResourceEmployment
    {
        public string EmploymentType { get; set; }

        public string EmploymentTypeDescription { get; set; }

        public bool MainPosition { get; set; }

        public Decimal Percentage { get; set; }

        public string PostId { get; set; }

        public string PostIdDescription { get; set; }

        public string PostCode { get; set; }

        public string PostCodeDescription { get; set; }

        public object SalaryPerYears { get; set; }

        public object Rates { get; set; }

        [XmlElement(DataType = "date")]
        public DateTime SeniorityDate { get; set; }

        public string WageRule { get; set; }

        [XmlArrayItem("Relation", IsNullable = false)]
        public ExportInfoResourceEmploymentRelation[] Relations { get; set; }

        [XmlElement(DataType = "date")]
        public DateTime DateFrom { get; set; }

        [XmlElement(DataType = "date")]
        public DateTime DateTo { get; set; }

        public DateTime LastUpdate { get; set; }

        public string SequenceRef { get; set; }

        public string SequenceNo { get; set; }

        public DateTime TotalLastUpdate { get; set; }

        [XmlAttribute]
        public string ResourceId { get; set; }
    }

    [XmlType(AnonymousType = true)]
    public class ExportInfoResourceEmploymentRelation
    {
        public string Value { get; set; }

        public string Value1 { get; set; }

        public string Description { get; set; }

        [XmlElement(DataType = "date")]
        public DateTime DateFrom { get; set; }

        [XmlElement(DataType = "date")]
        public DateTime DateTo { get; set; }

        public string SequenceRef { get; set; }

        public string SequenceNo { get; set; }

        public DateTime LastUpdate { get; set; }

        [XmlAttribute]
        public string Type { get; set; }

        [XmlAttribute]
        public string Name { get; set; }

        [XmlAttribute]
        public string ElementType { get; set; }
    }

    [XmlType(AnonymousType = true)]
    public class ExportInfoResourceRelation
    {
        public string Value { get; set; }

        public string Value1 { get; set; }

        public string Description { get; set; }

        [XmlElement(DataType = "date")]
        public DateTime DateFrom { get; set; }

        [XmlElement(DataType = "date")]
        public DateTime DateTo { get; set; }

        public DateTime LastUpdate { get; set; }

        [XmlAttribute]
        public string Type { get; set; }

        [XmlAttribute]
        public string Name { get; set; }

        [XmlAttribute]
        public string ElementType { get; set; }
    }
}
