using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Input;

namespace AcadExts
{
    public class DelegateCommand : ICommand
    {
        //#pragma warning disable 67
        public event EventHandler CanExecuteChanged;
        //#pragma warning restore 67

        private readonly Action _action;
        private readonly Predicate<object> _canExecute;

        public DelegateCommand(Action action, Predicate<object> canExecute)
        {
            if (action == null) { throw new ArgumentNullException("execute"); }
            _action = action;
            _canExecute = canExecute;
        }

        public DelegateCommand(Action action) : this(action, null)
        {
        }

        public void Execute(object Parameter)
        {
            _action();
        }
        public bool CanExecute(object parameter)
        {
            return (_canExecute == null) ? true : _canExecute(parameter);
        }

        public void RaiseCanExecuteChanged()
        {
            if (CanExecuteChanged != null)
            {
                CanExecuteChanged(this, new EventArgs());
            }
        }
    }
}
