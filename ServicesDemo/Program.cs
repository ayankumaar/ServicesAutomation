using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Threading;
using System.Windows.Automation;

namespace ServicesDemo
{
    class Program
    {
        static void Main(string[] args)
        {

            try
            {
                //Start the services       
                Process.Start("services.msc");
            }
            catch (Exception e)
            {
                Console.WriteLine("Couldn't start the process \"services.msc\" ");
                Console.WriteLine(e.Message.ToString());
                throw;
            }

            //TODO: Change the timer logic?? It takes around 4 seconds in my PC for the Services window to expose its Automation Properties.
            Thread.Sleep(7000);

            PropertyCondition typeCondition = new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Window);

            //Create a property condition with the application's name
            PropertyCondition nameCondition = new PropertyCondition(AutomationElement.NameProperty, "Services");

            //Create the conjunction condition
            AndCondition andCondition = new AndCondition(typeCondition, nameCondition);

            //Find the Main Window of Services as the root element.
            AutomationElement mainWnd = AutomationElement.RootElement.FindFirst(TreeScope.Children, andCondition);

            if (mainWnd == null)
            {
                //1.Log the exception

                //2. Then throw
                throw new InvalidOperationException("Couldn't find the Services MainWindow");
            }

            //The Automation id of the list which contains the services is "12786" 
            PropertyCondition findTheListById = new PropertyCondition(AutomationElement.AutomationIdProperty, "12786");
            AutomationElement getTheList = mainWnd.FindFirst(TreeScope.Descendants, findTheListById);


            if (getTheList == null)
            {
                //1.Log the exception

                //2. Then throw
                throw new InvalidOperationException("Couldn't find the list which contains Services with AutomationId \"12786\" ");
            }

            //The services present in the list are of LocalizedControlType "item"
            PropertyCondition cond = new PropertyCondition(AutomationElement.LocalizedControlTypeProperty, "item");

            //Find all the services
            AutomationElementCollection listItems = getTheList.FindAll(TreeScope.Descendants, cond);

            if (listItems == null)
            {
                //1.Log the exception

                //2.Then throw
                throw new InvalidOperationException("Can't find the services in the list");
            }

            if (listItems.Count == 0)
            {
                Console.WriteLine("No services found.");
                Thread.Sleep(1000);
                Environment.Exit(0);
            }

            Console.WriteLine("Total services found:" + " " + listItems.Count);

            //List to store the properties of all the Services
            List<ServiceDescription> listOfServiceReports = new List<ServiceDescription>();

            foreach (AutomationElement listItem in listItems)
            {
                //Find all the children of a service
                AutomationElementCollection allItems = listItem.FindAll(TreeScope.Children, Condition.TrueCondition);

                //TODO: What happens if no attributes for a service is found ??
                //if (allItems.Count == 0)
                //{
                //    throw new InvalidOperationException("No attribute found for the service:" + " " + listItem.Current.Name);
                //}

                ServiceDescription desc = new ServiceDescription();

                int x = 0;
                foreach (AutomationElement elem in allItems)
                {
                    /* The first three children of a service are "Name", "Description", "Status" followed by "StartUpType" and "LogOnAs" 
                    i.e. the same order in which the services window displays these properties
                    */
                    if (x < 3)
                    {
                        if (x == 0)
                            //Name of the Service
                            desc.name = elem.Current.Name;
                        else if (x == 1)
                            //Decription of the Service
                            desc.description = elem.Current.Name;
                        else if (x == 2)
                            //Status of the Service
                            desc.stauts = elem.Current.Name;
                    }
                    //Once done with the first three children, come out of the loop.
                    else
                    {
                        break;
                    }

                    x++;
                }

                //Add the properties of the service to a list
                listOfServiceReports.Add(desc);
            }

            try
            {
                Console.WriteLine("Trying to save the Service report in C drive in JSON Format");
                File.WriteAllText(@"c:\ServicesStatus.json", JsonConvert.SerializeObject(listOfServiceReports));
                Console.WriteLine("Please check the file \"ServicesStatus.json\" in C drive to see the details");
            }
            catch (Exception e)
            {
                Console.WriteLine("Couldn't save the file in C drive. Please check if there is any permission issue");
                Console.WriteLine(e.Message.ToString());
                Console.WriteLine("No worries ! I can print the output for you.");
                Thread.Sleep(2000);
                Console.WriteLine();
                foreach (var report in listOfServiceReports)
                {
                    Console.WriteLine(report.name);
                    Console.WriteLine(report.description);
                    Console.WriteLine(report.stauts);
                }
            }

            Console.ReadKey();
        }
    }
}
