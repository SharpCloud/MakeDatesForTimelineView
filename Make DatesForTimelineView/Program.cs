using System;
using System.Configuration;
using System.Security.Policy;
using SC.API.ComInterop;
using SC.API.ComInterop.Models;
using Attribute = SC.API.ComInterop.Models.Attribute;

/*
    Simple example use to show how items can be clipped to fit in a filtered time view.
 */ 


namespace Make_DatesForTimelineView
{
    class Program
    {
        static void Main(string[] args)
        {
            var userid = ConfigurationManager.AppSettings["userid"];
            if (string.IsNullOrEmpty(userid))
            {
                Console.WriteLine("Please enter your username");
                userid = Console.ReadLine();
            }
            var passwd = ConfigurationManager.AppSettings["passwd"];
            if (string.IsNullOrEmpty(passwd))
            {
                Console.WriteLine("Please enter your password");
                passwd = Console.ReadLine();
            }
            var URL = ConfigurationManager.AppSettings["URL"];
            var StartWindowDate = ConfigurationManager.AppSettings["StartWindowDate"];
            var EndWindowDate = ConfigurationManager.AppSettings["EndWindowDate"];
            var StartFieldName = ConfigurationManager.AppSettings["StartFieldName"];
            var EndFieldName = ConfigurationManager.AppSettings["EndFieldName"];

            var StoryCount = int.Parse(ConfigurationManager.AppSettings["StoryCount"]);

            var startDate = DateTime.Parse(StartWindowDate);
            var endDate = DateTime.Parse(EndWindowDate);

            int fixedStories = 0;
            try
            {
                var sc = new SharpCloudApi(userid, passwd);
                for (int i = 0; i < StoryCount; i++)
                {
                    var key = string.Format("Story{0:d2}", i + 1);
                    var storyid = ConfigurationManager.AppSettings[key];
                    if (!string.IsNullOrEmpty(storyid))
                    {
                        // fix this story
                        Console.WriteLine("Opening story {0}", i + 1);
                        var story = sc.LoadStory(storyid);
                        Console.WriteLine("Adjusting story {0}", story.Name);
                        FixStory(story, startDate, endDate, StartFieldName, EndFieldName);
                        story.Save();
                        Console.WriteLine("Saving story {0}", story.Name);

                        fixedStories++;
                    }
                }
                Console.WriteLine("{0} Stories Fixed.", fixedStories);
            }
            catch (Exception exception)
            {
                Console.WriteLine("There was an error.");
                Console.WriteLine(exception.ToString());
                throw;
            }
        }

        static void FixStory(Story story, DateTime startDate, DateTime endDate, string StartName, string EndName)
        {
            Attribute attSt = story.Attribute_FindByName(StartName); // ensure this exists
            if (attSt == null)
                attSt = story.Attribute_Add(StartName, Attribute.AttributeType.Date); 
            
            Attribute attEnd = story.Attribute_FindByName(EndName); // ensure this exists
            if (attEnd == null)
                attEnd = story.Attribute_Add(EndName, Attribute.AttributeType.Date); 

            foreach (var item in story.Items)
            {
                var st = item.StartDate;
                var end = st.AddDays(item.DurationInDays);

                if (end < startDate)
                {
                    // all of item before the window
                    item.RemoveAttributeValue(attSt);
                    item.RemoveAttributeValue(attEnd);
                }
                else if (st > endDate)
                {
                    // all of item after the window
                    item.RemoveAttributeValue(attSt);
                    item.RemoveAttributeValue(attEnd);
                }
                else
                {
                    // item is in window
                    // start
                    if (st < startDate)
                    {
                        // clip the start
                        item.SetAttributeValue(attSt, startDate);
                    }
                    else
                    {
                        // use the start date
                        item.SetAttributeValue(attSt, st);
                    }
                    
                    // end 
                    if (end > endDate)
                    {
                        // clip the end 
                        item.SetAttributeValue(attEnd, endDate);
                    }
                    else 
                    {
                        // sue the calculated end
                        item.SetAttributeValue(attEnd, end);
                    }


                
                }
            }

        }

    }
}
