﻿//------------------------------------------------------------------------------
// <auto-generated>
//     This code was generated by a tool.
//     Runtime Version:4.0.30319.42000
//
//     Changes to this file may cause incorrect behavior and will be lost if
//     the code is regenerated.
// </auto-generated>
//------------------------------------------------------------------------------

using System;

namespace LinkteraRobotics.Read.Range.Force.Activities.Properties
{
    /// <summary>
    ///   A strongly-typed resource class, for looking up localized strings, etc.
    /// </summary>
    // This class was auto-generated by the StronglyTypedResourceBuilder
    // class via a tool like ResGen or Visual Studio.
    // To add or remove a member, edit your .ResX file then rerun ResGen
    // with the /str option, or rebuild your VS project.
    [global::System.CodeDom.Compiler.GeneratedCodeAttribute("System.Resources.Tools.StronglyTypedResourceBuilder", "16.0.0.0")]
    [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
    [global::System.Runtime.CompilerServices.CompilerGeneratedAttribute()]
    public class Resources
    {

        private static global::System.Resources.ResourceManager resourceMan;

        private static global::System.Globalization.CultureInfo resourceCulture;

        [global::System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Performance", "CA1811:AvoidUncalledPrivateCode")]
        internal Resources()
        {
        }

        /// <summary>
        ///   Returns the cached ResourceManager instance used by this class.
        /// </summary>
        [global::System.ComponentModel.EditorBrowsableAttribute(global::System.ComponentModel.EditorBrowsableState.Advanced)]
        public static global::System.Resources.ResourceManager ResourceManager
        {
            get
            {
                if (object.ReferenceEquals(resourceMan, null))
                {
                    global::System.Resources.ResourceManager temp = new global::System.Resources.ResourceManager("LinkteraRobotics.Read.Range.Force.Activities.Properties.Resources", typeof(Resources).Assembly);
                    resourceMan = temp;
                }
                return resourceMan;
            }
        }

        /// <summary>
        ///   Overrides the current thread's CurrentUICulture property for all
        ///   resource lookups using this strongly typed resource class.
        /// </summary>
        [global::System.ComponentModel.EditorBrowsableAttribute(global::System.ComponentModel.EditorBrowsableState.Advanced)]
        public static global::System.Globalization.CultureInfo Culture
        {
            get
            {
                return resourceCulture;
            }
            set
            {
                resourceCulture = value;
            }
        }

        /// <summary>
        ///   Looks up a localized string similar to Common.
        /// </summary>
        public static string Common_Category
        {
            get
            {
                return ResourceManager.GetString("Common_Category", resourceCulture);
            }
        }

        /// <summary>
        ///   Looks up a localized string similar to If set, continue executing the remaining activities even if the current activity has failed..
        /// </summary>
        public static string ContinueOnError_Description
        {
            get
            {
                return ResourceManager.GetString("ContinueOnError_Description", resourceCulture);
            }
        }

        /// <summary>
        ///   Looks up a localized string similar to ContinueOnError.
        /// </summary>
        public static string ContinueOnError_DisplayName
        {
            get
            {
                return ResourceManager.GetString("ContinueOnError_DisplayName", resourceCulture);
            }
        }

        /// <summary>
        ///   Looks up a localized string similar to Do.
        /// </summary>
        public static string Do
        {
            get
            {
                return ResourceManager.GetString("Do", resourceCulture);
            }
        }

        /// <summary>
        ///   Looks up a localized string similar to Specifies the amount of time (in milliseconds) to wait for the activity to run before an error is thrown. The default value is 60000 (1 minute)..
        /// </summary>
        public static string Timeout_Description
        {
            get
            {
                return ResourceManager.GetString("Timeout_Description", resourceCulture);
            }
        }

        /// <summary>
        ///   Looks up a localized string similar to Timeout (milliseconds).
        /// </summary>
        public static string Timeout_DisplayName
        {
            get
            {
                return ResourceManager.GetString("Timeout_DisplayName", resourceCulture);
            }
        }

        /// <summary>
        ///   Looks up a localized string similar to The activity timed out and was canceled..
        /// </summary>
        public static string Timeout_Error
        {
            get
            {
                return ResourceManager.GetString("Timeout_Error", resourceCulture);
            }
        }
    }
}
