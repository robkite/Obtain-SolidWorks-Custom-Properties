using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using SolidWorks.Interop.swdocumentmgr;

namespace Obtain_SolidWorks_Custom_Properties {
    public static class Program {
        public static void Main() {

            Console.WriteLine("There are two methods for connecting to models and pulling out custom properties...");
            Console.WriteLine("  - ModelDoc2: access models that are open in the SolidWorks application");
            Console.WriteLine("  - SwDMDocument: access models using the document manager, this does not require utilisation of the SolidWorks application and is more performant");
            Console.WriteLine();

            Console.WriteLine("There are also two locations we can find custom properties within a model...");
            Console.WriteLine("  - Generic: generic custom properties are maintaining at the model level (part, assembly or drawing)");
            Console.WriteLine("  - Configuration Specific: there properties are unique to a given configuration within the model");
            Console.WriteLine();
            Console.WriteLine();


            Console.WriteLine("The source code provides example methods for...");
            Console.WriteLine("  - Obtaining child components from a root assembly (Document Manager example not included)");
            Console.WriteLine("  - Obtaining custom properties from models (generic or configuration specific)");

            Console.ReadLine();
        }

        #region Obtaining Components

        #region SolidWorks Application

         /// <summary>
        /// Obtains the components.
        /// </summary>
        /// <param name="swAssy">The SolidWorks assembly.</param>
        /// <returns>All child components.</returns>
        public static List<IComponent2> ObtainChildComponents(IAssemblyDoc swAssy) {
            
            // Obtain the root component
            var swModel = (IModelDoc2) swAssy;
            IConfiguration swConfiguration = swModel.GetActiveConfiguration();
            IComponent2 swRootComp = swConfiguration.GetRootComponent3(true);

            // Build a list of all child components
            return GetChildren(swRootComp);
        }

        /// <summary>
        /// Gets the children.
        /// </summary>
        /// <param name="component">The component.</param>
        /// <returns></returns>
        private static List<IComponent2> GetChildren(IComponent2 component) {
            Array childComponentsArray = component.GetChildren();

            var childComponents = new List<IComponent2>();
            if (childComponentsArray == null) return childComponents;
            foreach (IComponent2 comp in childComponentsArray) {
                childComponents.Add(comp); // Add the child
                childComponents.AddRange(GetChildren(comp)); // Recursively add children of the child
            }

            return childComponents;
        }

        #endregion

        #endregion


        #region Obtaining Custom Properties

        #region SolidWorks Application

        /// <summary>
        /// Gets the custom property.
        /// </summary>
        /// <param name="propertyName">Name of the property.</param>
        /// <param name="swModel">The SolidWorks model.</param>
        /// <param name="configuration">The configuration (the generic custom property will be used if no configuration specified).</param>
        /// <returns>The resolved property value, or null if the property is not found.</returns>
        public static string GetCustomProperty(string propertyName, IModelDoc2 swModel, string configuration = null) {
            
            // Obtain the custom property manager for the given configuration
            if (configuration == null) configuration = string.Empty;
            ICustomPropertyManager swCustomPropertyManager = swModel.Extension.CustomPropertyManager[configuration];
            if (swCustomPropertyManager == null) {
                Debug.Fail("Custom property manager is null, configuration " + configuration + " may not exist");
                return null;
            }

            // Obtain the custom property based on its name
            string valOut;
            string resolvedValOut;
            bool wasResolved;
            swCustomPropertyManager.Get5(propertyName, false, out valOut, out resolvedValOut, out wasResolved);

            // As the UseCache was set to false we will always receive a resolved value
            return resolvedValOut;
        }

        #endregion

        #region Document Manager

        /// <summary>
        /// Gets the custom property.
        /// </summary>
        /// <param name="propertyName">Name of the property.</param>
        /// <param name="swDmDocument">The SolidWorks document manager document.</param>
        /// <param name="configuration">The configuration (the generic custom property will be used if no configuration specified).</param>
        /// <returns>The property value, or null if the property is not found.</returns>
        public static string GetCustomProperty(string propertyName, ISwDMDocument swDmDocument, string configuration = null) {

            string propertyValue;
            SwDmCustomInfoType type;
            if (string.IsNullOrEmpty(configuration)) {
                // Get the generic (non configuration specific) custom property
                propertyValue = swDmDocument.GetCustomProperty(propertyName, out type);
            } else {
                // Get the configuration specific custom property
                ISwDMConfiguration swDmConfiguration = swDmDocument.ConfigurationManager.GetConfigurationByName(configuration);
                propertyValue = swDmConfiguration.GetCustomProperty(propertyName, out type);
            }

            return propertyValue;
        }
        
        #endregion

        #endregion
    }
}
