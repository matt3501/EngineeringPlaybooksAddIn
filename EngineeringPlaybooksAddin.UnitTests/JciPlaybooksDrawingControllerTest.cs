using System;
using EngineeringPlaybooksAddIn.Controllers;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace EngineeringPlaybooksAddin.UnitTests
{
    [TestClass]
    public class JciPlaybooksDrawingControllerTest : JciPlaybooksDrawingController
    {
        [TestMethod]
        public void ValidateAndTrimModelTest()
        {
            //Arrange
            var jsonText1 = "{\"title\":\"Johnson Controls Workflows Team Software Engineer\",\"description\":\"Software Engineer\",\"outcomes\":[{\"title\":\"Apply Domain Research\",\"contentUrl\":\"https://google.com\",\"childOutcomes\":[{\"title\":\"Outlining Problems\",\"contentUrl\":\"https://google.com\"},{\"title\":\"Identifying opportunities\",\"contentUrl\":\"https://google.com\"},{\"title\":\"Proposing solutions as direction\",\"contentUrl\":\"https://google.com\"},{\"title\":\"Broker with outside interests\",\"contentUrl\":\"https://google.com\"}]},{\"title\":\"Thoughtfully Communicate\",\"contentUrl\":\"https://google.com\",\"childOutcomes\":[{\"title\":\"Define Business and Functional Requirements\",\"contentUrl\":\"https://google.com\"},{\"title\":\"Define Non-Functional Requirements\",\"contentUrl\":\"https://google.com\"},{\"title\":\"Drive Discovery of new work to be done\",\"contentUrl\":\"https://google.com\"},{\"title\":\"Prompt UXUI with Requirements\",\"contentUrl\":\"https://google.com\"},{\"title\":\"Form acceptance criteria\",\"contentUrl\":\"http://enfocussolutions.com/acceptance-criteria-for-business-analysts/\"}]}]}";
            var jsonText2 = "{\"title\":\"Johnson Controls Workflows Team Software Engineer\",\"description\":\"Software Engineer\",\"outcomes\":[{\"title\":\"Apply Domain Research\",\"contentUrl\":\"https://google.com\",\"childOutcomes\":[{\"title\":\"Outlining Problems\",\"contentUrl\":\"https://google.com\"},{\"title\":\"Identifying opportunities\",\"contentUrl\":\"https://google.com\"},{\"title\":\"Proposing solutions as direction\",\"contentUrl\":\"https://google.com\"},{\"title\":\"Broker with outside interests\",\"contentUrl\":\"https://google.com\"}]},{\"title\":\"Thoughtfully Communicate\",\"contentUrl\":\"https://google.com\",\"childOutcomes\":[{\"title\":\"Define Business and Functional Requirements\",\"contentUrl\":\"https://google.com\"},{\"title\":\"Define Non-Functional Requirements\",\"contentUrl\":\"https://google.com\"},{\"title\":\"Drive Discovery of new work to be done\",\"contentUrl\":\"https://google.com\"},{\"title\":\"Prompt UXUI with Requirements\",\"contentUrl\":\"https://google.com\"},{\"title\":\"Form acceptance criteria\",\"contentUrl\":\"http://enfocussolutions.com/acceptance-criteria-for-business-analysts/\"}]}]}";

            //Act
            var result1 = JciPlaybooksDrawingController.ValidateAndTrimModel(jsonText1);
            var result2 = JciPlaybooksDrawingController.ValidateAndTrimModel(jsonText2);

            //Assert
            Assert.AreEqual(2, result1.outcomes.Count);
            Assert.AreEqual(2, result1.outcomes.Count);
        }
    }
}
