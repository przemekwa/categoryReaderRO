using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;

namespace categoryReaderRO
{
    internal class Row : IEquatable<Row>
    {
        public Row()
        {
        }

        public string Category { get; internal set; }
        public Guid Id { get; internal set; }
        public string WGR { get; internal set; }
        public List<string> WUGR { get; internal set; }

        public bool Equals([AllowNull] Row other)
        {
            return Category == other.Category;
        }
    }
}