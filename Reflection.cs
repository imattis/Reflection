using System;
using System.Reflection;

namespace ForTests
{
    public static class Reflection
    {
        #region Reflection
        /// <summary>
        /// Searches in dll class of specified type and params of constructor wirh specified number and types.
        /// </summary>
        /// <returns>Object that meets the conditions</returns>
        public static Object GetReflectedObject(string dllName, Type classType, params Object[] constructorParamsArray)
        {
            try
            {
                var reflectedDLL = Assembly.LoadFrom(dllName);  //Load specified dll

                object providerObject = getClass(reflectedDLL.GetTypes(), classType, constructorParamsArray);

                if (providerObject != null)
                    return providerObject;
                else
                    throw new ApplicationException("Cannot get object: There are no dll|class|construcotor...");
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        /// <summary>
        /// Searches in dll class of specified type and params of constructor wirh specified number and types.
        /// </summary>
        /// <returns>Object that meets the conditions</returns>
        public static Object GetReflectedObject(string dllName, string className, params Object[] constructorParamsArray)
        {
            try
            {
                var reflectedDLL = Assembly.LoadFrom(dllName);  //Load specified dll

                object providerObject = getClass(reflectedDLL.GetTypes(), className, constructorParamsArray);

                if (providerObject != null)
                    return providerObject;
                else
                    throw new ApplicationException("Cannot get object: There are no dll|class|construcotor...");
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        /// <summary>
        /// Searches for specified type and start GetSpecifiedConstructor.
        /// </summary>
        /// <returns>Object that meets the conditions else null</returns>
        private static Object getClass(Type[] allClassTypes, Type desiredType, params Object[] paramsArray)
        {
            object providerObject = null;

            foreach (var tempType in allClassTypes)
            {
                if (desiredType.IsAssignableFrom(tempType))
                {
                    //Get object which has constructor that fits the parameters
                    providerObject = getConstructor(tempType, paramsArray);

                    if (providerObject != null) return providerObject;
                }
            }

            return null;
        }

        /// <summary>
        /// Searches for specified type and start GetSpecifiedConstructor.
        /// </summary>
        /// <returns>Object that meets the conditions else null</returns>
        private static Object getClass(Type[] allClassTypes, string desiredClassName, params Object[] paramsArray)
        {
            object providerObject = null;

            foreach (var tempType in allClassTypes)
            {
                if (desiredClassName == tempType.Name)
                {
                    //Get object which has constructor that fits the parameters
                    providerObject = getConstructor(tempType, paramsArray);

                    if (providerObject != null) return providerObject;
                }
            }

            return null;
        }

        /// <summary>
        /// Searches for specified constructor and start ConstructorHasCorrectParams
        /// </summary>
        /// <returns>Object that meets the conditions else null</returns>
        private static Object getConstructor(Type classType, params Object[] paramsArray)
        {
            ConstructorInfo[] allClassConstructors = classType.GetConstructors();

            foreach (ConstructorInfo classConstructor in allClassConstructors)
            {
                ParameterInfo[] constructorParams = classConstructor.GetParameters();

                int optionParamsCount = getNonOptionalParamsCount(constructorParams);

                if ((paramsArray.Length >= constructorParams.Length - optionParamsCount && paramsArray.Length <= constructorParams.Length))
                {
                    if (constructorHasCorrectParams(constructorParams, paramsArray))    //Check for all parameters types coincide
                    {
                        if (paramsArray.Length < constructorParams.Length)  //If there are some optional params. Fill optional params with "Type.Missing"
                        {
                            int startPos = paramsArray.Length;
                            int count = constructorParams.Length - paramsArray.Length;

                            Array.Resize(ref paramsArray, paramsArray.Length + count);

                            for (int i = startPos; i < paramsArray.Length; i++)
                            {
                                paramsArray[i] = Type.Missing;
                            }
                        }

                        return classConstructor.Invoke(paramsArray);    //In any case invoke founded constructor
                    }
                }
            }

            return null;
        }

        /// <summary>
        /// Count non optional params
        /// </summary>
        /// <param name="constructorParams">All constructor params</param>
        /// <returns>Only non optional params</returns>
        private static int getNonOptionalParamsCount(ParameterInfo[] constructorParams)
        {
            int optionParamsCount = 0;

            foreach (var param in constructorParams)
            {
                if (param.IsOptional)
                    optionParamsCount++;
            }

            return optionParamsCount;
        }

        /// <summary>
        /// Checks for types of params of constructor. 
        /// </summary>
        /// <returns>Returns true if types are OK. And false if not.</returns>
        private static bool constructorHasCorrectParams(ParameterInfo[] constructorParams, params Object[] paramsArray)
        {
            for (int i = 0; i < constructorParams.Length; i++)
            {
                //Check for all parameters types coincide
                if (!constructorParams[i].IsOptional && constructorParams[i].ParameterType != paramsArray[i].GetType()) return false;
            }

            return true;
        }
        #endregion
    }
}
