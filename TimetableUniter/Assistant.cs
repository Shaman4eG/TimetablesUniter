using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TimetableUniter
{
    class Assistant
    {
        public int NumberOfShifts { get; set; } = 0;
        public string LastName { get; set; }

        private bool[] shiftPossible;

        private static readonly int maxShiftssInMonth = 62;

        public Assistant()
        {
            shiftPossible = new bool[maxShiftssInMonth];
        }

        public void AddShift(int index, bool newShift)
        {
            if (index < 0 && index >= shiftPossible.Length) throw new ArgumentOutOfRangeException();

            shiftPossible[index] = newShift;
        }

        public bool GetShift(int index)
        {
            if (index < 0 && index >= shiftPossible.Length) throw new ArgumentOutOfRangeException();

            return shiftPossible[index];
        }
    }
}
