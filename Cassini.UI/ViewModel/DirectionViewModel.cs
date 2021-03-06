﻿using System.Collections.ObjectModel;
using Cassini.Model;
using Prism.Events;
using Prism.Mvvm;

namespace Cassini.UI.ViewModel
{
    public class DirectionViewModel: BindableBase, ISelectableViewModel
    {

        private bool _isSelected;
        public System.Guid Guid { get; set; }
        public string Code { get; set; }
        public string Title { get; set; }

        public string FullName => $"{this.Code} {this.Title}";

        public bool IsSelected
        {
            get => _isSelected;
            set
            {
                if (_isSelected == value) return;
                _isSelected = value;
                RaisePropertyChanged(nameof(IsSelected));
            }
        }

    }
}