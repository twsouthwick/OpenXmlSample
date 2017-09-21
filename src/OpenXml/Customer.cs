// Licensed under the MIT license. See LICENSE file in the samples root for full license information.

using System;

namespace OpenXml.Models
{
    public class Customer
    {
        public string Name { get; set; }

        public string Address { get; set; }

        public string City { get; set; }

        public string State { get; set; }

        public string ZipCode { get; set; }

        public DateTimeOffset DateEntered { get; set; }
    }
}
