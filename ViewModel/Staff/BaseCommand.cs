using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Input;

namespace DBClasses.ViewModel.Staff
{
    public class BaseCommand : ICommand
    {
        private readonly Action command;
        private readonly Func<bool>? canExec;

        public event EventHandler? CanExecuteChanged;

        public BaseCommand(Action cmd, Func<bool>? canExec = null)
        {
            ArgumentNullException.ThrowIfNull(cmd, nameof(cmd));

            command = cmd;
            this.canExec = canExec;
        }

        public void NotifyCanExecuteChanged()
            => CanExecuteChanged?.Invoke(this, EventArgs.Empty);

        public bool CanExecute(object? parameter)
            => canExec is null || canExec.Invoke();

        public void Execute(object? parameter)
            => command.Invoke();
    }
}
