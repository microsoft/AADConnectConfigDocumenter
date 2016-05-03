//------------------------------------------------------------------------------------------------------------------------------------------
// <copyright file="LoggerCallContextItems.cs" company="Microsoft">
//      Copyright (c) Microsoft. All Rights Reserved.
//      Licensed under the MIT license. See LICENSE.txt file in the project root for full license information.
// </copyright>
// <summary>
// Logger Call Context Items class
// </summary>
//------------------------------------------------------------------------------------------------------------------------------------------

namespace AzureADConnectSyncDocumenter
{
    using System;
    using System.Collections;
    using System.Runtime.Remoting.Messaging;
    using System.Security;

    /// <summary>
    /// Provides methods to maintain a key/value dictionary that is stored in the <see cref="CallContext"/>.
    /// </summary>
    /// <remarks>
    /// A context item represents a key/value that needs to be logged with each message
    /// on the same CallContext.
    /// </remarks>
    internal static class LoggerCallContextItems
    {
        /// <summary>
        /// The name of the data slot in the <see cref="CallContext"/> used by the application block.
        /// </summary>
        private const string CallContextSlotName = "AzureADSyncDocumenter.ContextItems";

        /// <summary>
        /// Adds a key/value pair to a dictionary in the <see cref="CallContext"/>.  
        /// Each context item is recorded with every log entry.
        /// </summary>
        /// <param name="key">Hashtable key.</param>
        /// <param name="value">Value of the context item.  Byte arrays will be base64 encoded.</param>
        /// <example>The following example demonstrates use of the AddContextItem method.
        /// <code>Logger.SetContextItem("SessionID", myComponent.SessionId);</code></example>
        [SecurityCritical]
        public static void SetContextItem(object key, object value)
        {
            if (key == null)
            {
                throw new ArgumentNullException("key");
            }

            var contextItems = (Hashtable)CallContext.GetData(LoggerCallContextItems.CallContextSlotName) ?? new Hashtable();

            contextItems[key] = value;

            CallContext.SetData(LoggerCallContextItems.CallContextSlotName, contextItems);
        }

        /// <summary>
        /// Clears the context item.
        /// </summary>
        /// <param name="key">The key.</param>
        [SecurityCritical]
        public static void ClearContextItem(object key)
        {
            if (key == null)
            {
                throw new ArgumentNullException("key");
            }

            var contextItems = (Hashtable)CallContext.GetData(LoggerCallContextItems.CallContextSlotName) ?? new Hashtable();

            if (contextItems.ContainsKey(key))
            {
                contextItems.Remove(key);
            }

            CallContext.SetData(LoggerCallContextItems.CallContextSlotName, contextItems);
        }

        /// <summary>
        /// Empties the context items dictionary.
        /// </summary>
        [SecurityCritical]
        public static void FlushContextItems()
        {
            CallContext.FreeNamedDataSlot(LoggerCallContextItems.CallContextSlotName);
        }

        /// <summary>
        /// Gets the context items.
        /// </summary>
        /// <returns>The Hashtable of items on the call context.</returns>
        [SecurityCritical]
        public static Hashtable GetContextItems()
        {
            return (Hashtable)CallContext.GetData(LoggerCallContextItems.CallContextSlotName);
        }

        /// <summary>
        /// Gets the context item value.
        /// </summary>
        /// <param name="contextData">The context data.</param>
        /// <returns>The context item value.</returns>
        public static string GetContextItemValue(object contextData)
        {
            var value = string.Empty;

            if (contextData != null)
            {
                value = contextData.GetType() == typeof(byte[]) ? Convert.ToBase64String((byte[])contextData) : contextData.ToString();
            }

            return value;
        }
    }
}
