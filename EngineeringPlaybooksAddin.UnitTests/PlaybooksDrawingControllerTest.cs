using System;
using System.IO;
using EngineeringPlaybooksAddIn.Controllers;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace EngineeringPlaybooksAddin.UnitTests
{
    [TestClass]
    public class PlaybooksDrawingControllerTest : PlaybooksDrawingController
    {
        [TestMethod]
        public void ValidateAndTrimModelTest()
        {
            //Arrange
            var jsonText1 = "{\"title\":\"Software Engineer\",\"description\":\"Software Engineer\",\"outcomes\":[{\"title\":\"Apply Domain Research\",\"contentUrl\":\"https://google.com\",\"childOutcomes\":[{\"title\":\"Outlining Problems\",\"contentUrl\":\"https://google.com\"},{\"title\":\"Identifying opportunities\",\"contentUrl\":\"https://google.com\"},{\"title\":\"Proposing solutions as direction\",\"contentUrl\":\"https://google.com\"},{\"title\":\"Broker with outside interests\",\"contentUrl\":\"https://google.com\"}]},{\"title\":\"Thoughtfully Communicate\",\"contentUrl\":\"https://google.com\",\"childOutcomes\":[{\"title\":\"Define Business and Functional Requirements\",\"contentUrl\":\"https://google.com\"},{\"title\":\"Define Non-Functional Requirements\",\"contentUrl\":\"https://google.com\"},{\"title\":\"Drive Discovery of new work to be done\",\"contentUrl\":\"https://google.com\"},{\"title\":\"Prompt UXUI with Requirements\",\"contentUrl\":\"https://google.com\"},{\"title\":\"Form acceptance criteria\",\"contentUrl\":\"http://enfocussolutions.com/acceptance-criteria-for-business-analysts/\"}]}]}";
            
            //Act
            var result1 = ValidateAndTrimModel(jsonText1);

            //Assert
            Assert.AreEqual(2, result1.outcomes.Count);
        }
        
        [TestMethod]
        public void GetOffsetAngleRadiansTest_4_1()
        {
            //Arrange
            var parentPath = Path.GetDirectoryName(Environment.CurrentDirectory);
            var jsonText1 = File.ReadAllText(Path.Combine(parentPath ?? "", @"..\..\EngineeringPlaybooksAddIn\samples\knowledge_4_1.json"));

            //Act
            var validatedAndTrimmedModel = ValidateAndTrimModel(jsonText1);
            
            var outcome0 = validatedAndTrimmedModel.outcomes[0];

            var ellipseVectors = GetEllipseVertices(validatedAndTrimmedModel);

            var tangent0 = GetOffsetAngleRadians(outcome0, ellipseVectors[0]);

            //Assert
            Assert.AreEqual(4, validatedAndTrimmedModel.outcomes.Count);

            Assert.AreEqual(Math.PI / -4, tangent0, 0.01);
        }

        [TestMethod]
        public void GetOffsetAngleRadiansTest_1_2()
        {
            //Arrange
            var parentPath = Path.GetDirectoryName(Environment.CurrentDirectory);
            var jsonText1 = File.ReadAllText(Path.Combine(parentPath ?? "", @"..\..\EngineeringPlaybooksAddIn\samples\knowledge_1_2.json"));

            //Act
            var validatedAndTrimmedModel = ValidateAndTrimModel(jsonText1);
            
            var outcome0 = validatedAndTrimmedModel.outcomes[0];

            var ellipseVectors = GetEllipseVertices(validatedAndTrimmedModel);

            var tangent0 = GetOffsetAngleRadians(outcome0, ellipseVectors[0]);

            //Assert
            Assert.AreEqual(1, validatedAndTrimmedModel.outcomes.Count);

            Assert.AreEqual(Math.PI, tangent0, 0.001);
        }

        [TestMethod]
        public void GetOffsetAngleRadiansTest_1_3()
        {
            //Arrange
            var parentPath = Path.GetDirectoryName(Environment.CurrentDirectory);
            var jsonText1 = File.ReadAllText(Path.Combine(parentPath ?? "", @"..\..\EngineeringPlaybooksAddIn\samples\knowledge_1_3.json"));

            //Act
            var validatedAndTrimmedModel = ValidateAndTrimModel(jsonText1);
            
            var outcome0 = validatedAndTrimmedModel.outcomes[0];

            var ellipseVectors = GetEllipseVertices(validatedAndTrimmedModel);

            var tangent0 = GetOffsetAngleRadians(outcome0, ellipseVectors[0]);

            //Assert
            Assert.AreEqual(1, validatedAndTrimmedModel.outcomes.Count);

            Assert.AreEqual(Math.PI / 2.0, tangent0, 0.001);
        }
    }
}
