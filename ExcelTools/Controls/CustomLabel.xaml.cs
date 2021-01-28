using System.Windows;
using System.Windows.Controls;

namespace ExcelTools.Controls
{

    public partial class CustomLabel : UserControl
    {
        public string Header
        {
            get => (string)this.GetValue(HeaderProperty);
            set => this.SetValue(HeaderProperty, value);
        }

        public static readonly DependencyProperty HeaderProperty
            = DependencyProperty.Register(
                nameof(Header),
                typeof(string),
                typeof(CustomLabel),
                new PropertyMetadata(string.Empty)
            );

        public string SubHeader
        {
            get => (string)this.GetValue(SubHeaderProperty);
            set => this.SetValue(SubHeaderProperty, value);
        }

        public static readonly DependencyProperty SubHeaderProperty
            = DependencyProperty.Register(
                nameof(SubHeader),
                typeof(string),
                typeof(CustomLabel),
                new PropertyMetadata(string.Empty)
            );

        public CustomLabel()
        {
            this.InitializeComponent();

            this.DataContext = this;
        }
    }
}
