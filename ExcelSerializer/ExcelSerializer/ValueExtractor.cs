using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace ExcelSerializer
{
    public class ValueExtractor
    {
        public Dictionary<string, ObjectWrapper> DirectValues { get; set; }
        public Dictionary<string, ObjectWrapper> Objects { get; set; }
        public Dictionary<string, ObjectWrapper> ObjectLists { get; set; }

        public ValueExtractor ()
        {
            DirectValues = new Dictionary<string, ObjectWrapper>();
            Objects = new Dictionary<string, ObjectWrapper>();
            ObjectLists = new Dictionary<string, ObjectWrapper>();
        }

        public void Extract ( object instance )
        {
            if ( instance == null ) return;

            IList<PropertyInfo> props = new List<PropertyInfo>(instance.GetType().GetProperties());

            foreach ( PropertyInfo prop in props )
            {
                var t = prop.PropertyType;
                // direct value
                if ( t.IsPrimitive || t.IsValueType || ( t == typeof(string) ) )
                {
                    DirectValues.Add(prop.Name, new ObjectWrapper
                    {
                        ObjectType = ObjectTypes.DirectValue,
                        Value = prop.GetValue(instance)
                    });
                }
                else if ( t.GetInterface("ICollection") != null )
                {
                    var collection = prop.GetValue(instance) as ICollection;

                    //bool isValid = collection.Select(x => ((dynamic)x).GetType()).Distinct().Count() == 1;
                    //if ( !isValid ) continue;

                    var objectWrapper = new ObjectWrapper
                    {
                        ObjectType = ObjectTypes.ObjectList,
                        Value = new List<object>()
                    };

                    var ve = new ValueExtractor();

                    foreach ( var item in collection )
                    {
                        ( objectWrapper.Value as List<object> ).Add(item);
                    }

                    ObjectLists.Add(prop.Name, objectWrapper);
                }
                else if ( t.GetInterface("IEnumerable") != null )
                {
                    var enumerable = prop.GetValue(instance) as IEnumerable<object>;
                    var isValid = enumerable.Select(x => x.GetType()).Distinct().Count() == 1;

                    if ( !isValid ) continue;

                    Console.WriteLine(prop.Name);

                    //var ve = new ValueExtractor();

                    //foreach ( var item in enumerable )
                    //{
                    //    var xs = new ValueExtractor();
                    //    xs.Extract(prop.GetValue(item));
                    //    Objects.Add(prop.Name, new ObjectWrapper
                    //    {
                    //        ObjectType = ObjectTypes.Object,
                    //        Value = xs.DirectValues
                    //    });
                    //}
                }
                // object
                else
                {
                    var xs = new ValueExtractor();
                    xs.Extract(prop.GetValue(instance));
                    Objects.Add(prop.Name, new ObjectWrapper
                    {
                        ObjectType = ObjectTypes.Object,
                        Value = xs.DirectValues
                    });
                }
            }
        }

    }

    public class ObjectWrapper
    {
        public object Value { get; set; }
        public ObjectTypes ObjectType { get; set; }
    }

    public enum ObjectTypes
    {
        DirectValue,
        ObjectList,
        Object,
        DirectValueList
    }
}