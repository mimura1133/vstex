/********************************************************************************************

Copyright (c) Microsoft Corporation 
All rights reserved. 

Microsoft Public License: 

This license governs use of the accompanying software. If you use the software, you 
accept this license. If you do not accept the license, do not use the software. 

1. Definitions 
The terms "reproduce," "reproduction," "derivative works," and "distribution" have the 
same meaning here as under U.S. copyright law. 
A "contribution" is the original software, or any additions or changes to the software. 
A "contributor" is any person that distributes its contribution under this license. 
"Licensed patents" are a contributor's patent claims that read directly on its contribution. 

2. Grant of Rights 
(A) Copyright Grant- Subject to the terms of this license, including the license conditions 
and limitations in section 3, each contributor grants you a non-exclusive, worldwide, 
royalty-free copyright license to reproduce its contribution, prepare derivative works of 
its contribution, and distribute its contribution or any derivative works that you create. 
(B) Patent Grant- Subject to the terms of this license, including the license conditions 
and limitations in section 3, each contributor grants you a non-exclusive, worldwide, 
royalty-free license under its licensed patents to make, have made, use, sell, offer for 
sale, import, and/or otherwise dispose of its contribution in the software or derivative 
works of the contribution in the software. 

3. Conditions and Limitations 
(A) No Trademark License- This license does not grant you rights to use any contributors' 
name, logo, or trademarks. 
(B) If you bring a patent claim against any contributor over patents that you claim are 
infringed by the software, your patent license from such contributor to the software ends 
automatically. 
(C) If you distribute any portion of the software, you must retain all copyright, patent, 
trademark, and attribution notices that are present in the software. 
(D) If you distribute any portion of the software in source code form, you may do so only 
under this license by including a complete copy of this license with your distribution. 
If you distribute any portion of the software in compiled or object code form, you may only 
do so under a license that complies with this license. 
(E) The software is licensed "as-is." You bear the risk of using it. The contributors give 
no express warranties, guarantees or conditions. You may have additional consumer rights 
under your local laws which this license cannot change. To the extent permitted under your 
local laws, the contributors exclude the implied warranties of merchantability, fitness for 
a particular purpose and non-infringement.

********************************************************************************************/

using System;
using System.Collections;
using System.ComponentModel;

namespace VsTeXProject.VisualStudio.Project
{
    /// <summary>
    ///     The purpose of DesignPropertyDescriptor is to allow us to customize the
    ///     display name of the property in the property grid.  None of the CLR
    ///     implementations of PropertyDescriptor allow you to change the DisplayName.
    /// </summary>
    public class DesignPropertyDescriptor : PropertyDescriptor
    {
        private TypeConverter converter;
        private readonly Hashtable editors = new Hashtable(); // Type -> editor instance
        private readonly PropertyDescriptor property; // Base property descriptor

        /// <summary>
        ///     Constructor.  Copy the base property descriptor and also hold a pointer
        ///     to it for calling its overridden abstract methods.
        /// </summary>
        public DesignPropertyDescriptor(PropertyDescriptor prop)
            : base(prop)
        {
            if (prop == null)
            {
                throw new ArgumentNullException("prop");
            }

            property = prop;

            var attr = prop.Attributes[typeof (DisplayNameAttribute)] as DisplayNameAttribute;

            if (attr != null)
            {
                DisplayName = attr.DisplayName;
            }
            else
            {
                DisplayName = prop.Name;
            }
        }


        /// <summary>
        ///     Delegates to base.
        /// </summary>
        public override string DisplayName { get; }

        /// <summary>
        ///     Delegates to base.
        /// </summary>
        public override Type ComponentType
        {
            get { return property.ComponentType; }
        }

        /// <summary>
        ///     Delegates to base.
        /// </summary>
        public override bool IsReadOnly
        {
            get { return property.IsReadOnly; }
        }

        /// <summary>
        ///     Delegates to base.
        /// </summary>
        public override Type PropertyType
        {
            get { return property.PropertyType; }
        }


        /// <summary>
        ///     Return type converter for property
        /// </summary>
        public override TypeConverter Converter
        {
            get
            {
                if (converter == null)
                {
                    var attr =
                        (PropertyPageTypeConverterAttribute) Attributes[typeof (PropertyPageTypeConverterAttribute)];
                    if (attr != null && attr.ConverterType != null)
                    {
                        converter = (TypeConverter) CreateInstance(attr.ConverterType);
                    }

                    if (converter == null)
                    {
                        converter = TypeDescriptor.GetConverter(PropertyType);
                    }
                }
                return converter;
            }
        }


        /// <summary>
        ///     Delegates to base.
        /// </summary>
        public override object GetEditor(Type editorBaseType)
        {
            var editor = editors[editorBaseType];
            if (editor == null)
            {
                for (var i = 0; i < Attributes.Count; i++)
                {
                    var attr = Attributes[i] as EditorAttribute;
                    if (attr == null)
                    {
                        continue;
                    }
                    var editorType = Type.GetType(attr.EditorBaseTypeName);
                    if (editorBaseType == editorType)
                    {
                        var type = GetTypeFromNameProperty(attr.EditorTypeName);
                        if (type != null)
                        {
                            editor = CreateInstance(type);
                            editors[type] = editor; // cache it
                            break;
                        }
                    }
                }
            }
            return editor;
        }


        /// <summary>
        ///     Convert name to a Type object.
        /// </summary>
        public virtual Type GetTypeFromNameProperty(string typeName)
        {
            return Type.GetType(typeName);
        }


        /// <summary>
        ///     Delegates to base.
        /// </summary>
        public override bool CanResetValue(object component)
        {
            var result = property.CanResetValue(component);
            return result;
        }

        /// <summary>
        ///     Delegates to base.
        /// </summary>
        public override object GetValue(object component)
        {
            var value = property.GetValue(component);
            return value;
        }

        /// <summary>
        ///     Delegates to base.
        /// </summary>
        public override void ResetValue(object component)
        {
            property.ResetValue(component);
        }

        /// <summary>
        ///     Delegates to base.
        /// </summary>
        public override void SetValue(object component, object value)
        {
            property.SetValue(component, value);
        }

        /// <summary>
        ///     Delegates to base.
        /// </summary>
        public override bool ShouldSerializeValue(object component)
        {
            var result = property.ShouldSerializeValue(component);
            return result;
        }
    }
}