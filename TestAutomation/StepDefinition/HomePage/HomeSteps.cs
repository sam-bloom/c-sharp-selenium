using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using TestAutomation.FrameworkSupports;
using TechTalk.SpecFlow;
using TestAutomation.PageObjectModel.GoogleHome;

namespace TestAutomation.StepDefinition.HomePage
{
    [Binding]
    public class HomeSteps : GeneralMethods
    {

        DriverScript dr = new DriverScript();
        GoogleHomePage Home = new GoogleHomePage();

        [Given(@"I launch Google Test application")]
        public void GivenILaunchApplication()
        {
            BaseUtilities.applicationType = "web";
            BaseUtilities.GetInitialConfigData();
            dr.InitializeDriver();
            Home.launchApplication();
        }

  
    }
}
