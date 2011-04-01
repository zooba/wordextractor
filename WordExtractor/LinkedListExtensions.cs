using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace WordExtractor
{
    static class LinkedListExtensions
    {
        public static LinkedListNode<T> AddAfter<T>(this LinkedList<T> list, LinkedListNode<T> node, LinkedListNode<T> from, LinkedListNode<T> to)
        {
            if (list == null) throw new ArgumentNullException("list");

            var at = node;
            if (at == null) at = list.AddFirst(default(T));
            for (var c = from; c != null && c != to.Next; c = c.Next)
            {
                at = list.AddAfter(at, c.Value);
            }
            if (node == null) list.RemoveFirst();
            return at;
        }

        public static void Remove<T>(this LinkedListNode<T> node)
        {
            if (node == null) throw new ArgumentNullException("node");
            node.List.Remove(node);
        }

        public static void RemoveTo<T>(this LinkedListNode<T> start, LinkedListNode<T> end)
        {
            if (start == null) throw new ArgumentNullException("start");
            var list = start.List;

            while (start.Next != end) list.Remove(start.Next);
            list.Remove(start);
            if (end != null) list.Remove(end);
        }
    }
}
