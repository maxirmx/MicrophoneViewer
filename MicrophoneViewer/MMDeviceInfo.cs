using NAudio.CoreAudioApi;
using NAudio.CoreAudioApi.Interfaces;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Runtime.InteropServices;

namespace MicrophoneViewer
{
    public class MMDeviceInfo : ICustomTypeDescriptor
    {
        [Category("General")]
        public string ID { get; }
        [Category("General")]
        public string FriendlyName { get; }
        [Category("General")]
        public string DeviceFriendlyName { get; }
        [Category("General")]
        public string Description { get; }
        [Category("General")]
        public string Manufacturer { get; }
        [Category("General")]
        public string State { get; }
        [Category("Audio")]
        public int Channels { get; }
        [Category("Audio")]
        public int SampleRate { get; }
        [Category("Audio")]
        public int BitsPerSample { get; }
        [Category("Audio")]
        public float Volume { get; }
        [Category("Audio")]
        public bool Mute { get; }

        private readonly Dictionary<string, object> _properties;

        public MMDeviceInfo(MMDevice device)
        {
            ID = device.ID;
            FriendlyName = device.FriendlyName;
            DeviceFriendlyName = device.DeviceFriendlyName;
            Description = device.DeviceFriendlyName;
            Manufacturer = GetProperty(device, "System.Manufacturer");
            State = device.State.ToString();
            Channels = device.AudioClient.MixFormat.Channels;
            SampleRate = device.AudioClient.MixFormat.SampleRate;
            BitsPerSample = device.AudioClient.MixFormat.BitsPerSample;
            Volume = device.AudioEndpointVolume?.MasterVolumeLevelScalar ?? 0;
            Mute = device.AudioEndpointVolume?.Mute ?? false;

            _properties = new Dictionary<string, object>();
            for (int i = 0; i < device.Properties.Count; i++)
            {
                var key = device.Properties.Get(i);
                var value = device.Properties.GetValue(i);
                _properties[GetFriendlyPropertyKeyName(key)] = ExtractPropVariantValue(value);
            }
        }

        private string GetFriendlyPropertyKeyName(PropertyKey propertyKey)
        {
            // Dictionary of well-known property keys
            var wellKnownKeys = new Dictionary<string, string>
            {
                // Device properties
                {"{a45c254e-df1c-4efd-8020-67d146a850e0},2", "Device Name"},
                {"{a45c254e-df1c-4efd-8020-67d146a850e0},14", "Friendly Name"},
                {"{a45c254e-df1c-4efd-8020-67d146a850e0},13", "Manufacturer"},
                {"{a45c254e-df1c-4efd-8020-67d146a850e0},15", "Description"},
                {"{a45c254e-df1c-4efd-8020-67d146a850e0},24", "Enumerator Name"},
                {"{b3f8fa53-0004-438e-9003-51a46e139bfc},2", "Device Description"},
                {"{83da6326-97a6-4088-9453-a1923f573b29},3", "Icon Path"},
                
                // Audio endpoint properties
                {"{1da5d803-d492-4edd-8c23-e0c0ffee7f0e},1", "Form Factor"},
                {"{1da5d803-d492-4edd-8c23-e0c0ffee7f0e},2", "Spatial Audio"},
                {"{f19f064d-082c-4e27-bc73-6882a1bb8e4c},0", "Default Format"},
                
                // FX properties
                {"{d04e05a6-594b-4fb6-a80d-01af5eed7d1d},1", "APO Effects"},
                {"{5f8ec2c5-fd66-4148-95a6-2d470477d7a1},4", "Signal Processing Mode"},
                
                // Common property keys
                {"{def99636-439f-4c71-8f6a-57ee0470707d},3", "Driver Provider"},
                {"{78c34fc8-104a-4aca-9ea4-524d52996e57},57", "Device Instance ID"}
            };

            Guid guid = propertyKey.formatId; // or propertyKey.FormatId
            int propertyId = propertyKey.propertyId; // or propertyKey.PropertyId

            string keyString = $"{{{guid}}},{propertyId}";
            if (wellKnownKeys.TryGetValue(keyString, out string friendlyName))
            {
                return friendlyName;
            }

            return keyString;
        }

        private object ExtractPropVariantValue(object value)
        {
            // Handle PropVariant objects
            if (value is PropVariant propVariant)
            {
                try
                {
                    switch (propVariant.DataType)
                    {
                        case VarEnum.VT_LPWSTR:
                            return "LPWSTR: " + propVariant.Value;
                        case VarEnum.VT_UI4:
                            return "UI8: " + propVariant.Value;
                        case VarEnum.VT_I4:
                            return "I4: " + propVariant.Value;
                        case VarEnum.VT_UI8:
                            return "UI8: " + propVariant.Value;
                        case VarEnum.VT_I8:
                            return "I8: " + propVariant.Value;
                        case VarEnum.VT_UI2:
                            return "UI2: " + propVariant.Value;
                        case VarEnum.VT_I2:
                            return "I2: " + propVariant.Value;
                        case VarEnum.VT_UI1:
                            return "UI1: " + propVariant.Value;
                        case VarEnum.VT_I1:
                            return "I1: " + propVariant.Value;
                        case VarEnum.VT_BOOL:
                            return "BOOL: " + propVariant.Value;
                        case VarEnum.VT_R4:
                            return "R4: " + propVariant.Value;
                        case VarEnum.VT_R8:
                            return "R8L: " + propVariant.Value;
                        case VarEnum.VT_CLSID:
                            return "GUID: " + propVariant.Value;
                        case VarEnum.VT_BLOB:
                            return "[Binary data]";
                        case VarEnum.VT_EMPTY:
                            return "(empty)";
                        case VarEnum.VT_NULL:
                            return "(null)";
                        default:
                            return $"[PropVariant: {propVariant.DataType}]";
                    }
                }
                catch (Exception ex)
                {
                    return $"[Error extracting value: {ex.Message}]";
                }
            }

            return value?.ToString() ?? "(null)";
        }

        private string GetProperty(MMDevice device, string key)
        {
            for (int i = 0; i < device.Properties.Count; i++)
            {
                var propertyKey = device.Properties.Get(i);
                if (propertyKey.ToString().Contains(key))
                {
                    var value = device.Properties.GetValue(i);
                    return ExtractPropVariantValue(value)?.ToString() ?? string.Empty;
                }
            }
            return string.Empty;
        }

        #region ICustomTypeDescriptor Implementation

        public AttributeCollection GetAttributes()
        {
            return TypeDescriptor.GetAttributes(this, true);
        }

        public string GetClassName()
        {
            return TypeDescriptor.GetClassName(this, true);
        }

        public string GetComponentName()
        {
            return TypeDescriptor.GetComponentName(this, true);
        }

        public TypeConverter GetConverter()
        {
            return TypeDescriptor.GetConverter(this, true);
        }

        public EventDescriptor GetDefaultEvent()
        {
            return TypeDescriptor.GetDefaultEvent(this, true);
        }

        public PropertyDescriptor GetDefaultProperty()
        {
            return TypeDescriptor.GetDefaultProperty(this, true);
        }

        public object GetEditor(Type editorBaseType)
        {
            return TypeDescriptor.GetEditor(this, editorBaseType, true);
        }

        public EventDescriptorCollection GetEvents()
        {
            return TypeDescriptor.GetEvents(this, true);
        }

        public EventDescriptorCollection GetEvents(Attribute[] attributes)
        {
            return TypeDescriptor.GetEvents(this, attributes, true);
        }

        public PropertyDescriptorCollection GetProperties()
        {
            // Get the standard properties
            PropertyDescriptorCollection originalProperties =
                TypeDescriptor.GetProperties(this, true);

            // Create a list to hold all property descriptors
            List<PropertyDescriptor> allProperties = new List<PropertyDescriptor>();

            // Add the standard properties (excluding the Properties dictionary)
            foreach (PropertyDescriptor prop in originalProperties)
            {
                if (prop.Name != "Properties" && prop.Name != "_properties")
                {
                    allProperties.Add(prop);
                }
            }

            // Add each dictionary entry as its own property descriptor with decoded names
            foreach (KeyValuePair<string, object> kvp in _properties)
            {
                // The key should already be friendly, but ensure no duplicates by adding a counter if needed
                string propertyName = kvp.Key;
                int counter = 1;

                // Make sure property names are unique and valid
                while (allProperties.Any(p => p.Name == propertyName))
                {
                    propertyName = $"{kvp.Key} ({counter++})";
                }

                allProperties.Add(new DictionaryPropertyDescriptor(propertyName, kvp.Value));
            }

            return new PropertyDescriptorCollection(allProperties.ToArray());
        }

        public PropertyDescriptorCollection GetProperties(Attribute[] attributes)
        {
            return GetProperties();
        }

        public object GetPropertyOwner(PropertyDescriptor pd)
        {
            return this;
        }

        #endregion

        /// <summary>
        /// Custom PropertyDescriptor for dictionary entries
        /// </summary>
        private class DictionaryPropertyDescriptor : PropertyDescriptor
        {
            private readonly string _key;
            private readonly object _value;
            private readonly string _displayName;

            public DictionaryPropertyDescriptor(string key, object value)
                : base(MakeValidPropertyName(key), new Attribute[] { new CategoryAttribute("Device Properties") })
            {
                _key = key;
                _value = value;
                _displayName = key; // Store the original key as display name
            }

            public override string DisplayName => _displayName;

            public override Type ComponentType => typeof(MMDeviceInfo);

            public override bool IsReadOnly => true;

            public override Type PropertyType => _value?.GetType() ?? typeof(string);

            public override bool CanResetValue(object component) => false;

            public override object GetValue(object component) => _value;

            public override void ResetValue(object component) { }

            public override void SetValue(object component, object value) { }

            public override bool ShouldSerializeValue(object component) => false;

            // Helper method to ensure property names are valid identifiers
            private static string MakeValidPropertyName(string name)
            {
                if (string.IsNullOrEmpty(name))
                    return "Property";

                // Remove invalid characters
                string validName = new string(name.Select(c =>
                    char.IsLetterOrDigit(c) || c == '_' ? c : '_').ToArray());

                // Ensure it starts with a letter or underscore
                if (!char.IsLetter(validName[0]) && validName[0] != '_')
                    validName = "_" + validName;

                return validName;
            }
        }
    }
}
