using System;
using System.Collections.Generic;

namespace XlToDb
{
    public class Stack<T>
    {
        private readonly List<T> _stack;

        public Stack()
        {
            _stack = new List<T>();
        }

        public void Push(T obj)
        {
            _stack.Add(obj);
        }

        public object Pop()
        {
            int depth = _stack.Count - 1;
            var popped = _stack[depth];
            _stack.RemoveAt(depth);
            return popped;
        }

        public void Clear()
        {
            _stack.Clear();
        }

        public int Depth()
        {
            return _stack.Count;
        }
    }
}
