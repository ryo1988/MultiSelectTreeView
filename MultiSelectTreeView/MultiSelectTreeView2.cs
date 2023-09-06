using System;
using System.Collections.Generic;
using System.Collections.Immutable;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Media3D;

namespace MultiSelectTreeView
{
    public class MultiSelectTreeView2 : TreeView
    {
        public MultiSelectTreeView2()
        {
            GotFocus += OnTreeViewItemGotFocus;
            PreviewMouseLeftButtonDown += OnTreeViewItemPreviewMouseDown;
            PreviewMouseLeftButtonUp += OnTreeViewItemPreviewMouseUp;
        }

        private static TreeViewItem _selectTreeViewItemOnMouseUp;

        private ImmutableHashSet<object> _selectedItemsHashSet = ImmutableHashSet<object>.Empty;
        private object _startItem;

        public static readonly DependencyProperty IsItemSelectedProperty = DependencyProperty.RegisterAttached("IsItemSelected", typeof(Boolean), typeof(MultiSelectTreeView2));

        public static bool GetIsItemSelected(TreeViewItem element)
        {
            return (bool)element.GetValue(IsItemSelectedProperty);
        }

        public static void SetIsItemSelected(TreeViewItem element, Boolean value)
        {
            if (element == null) return;

            element.SetValue(IsItemSelectedProperty, value);
        }

        public static readonly DependencyProperty SelectedItemsProperty = DependencyProperty.RegisterAttached("SelectedItems", typeof(Array), typeof(MultiSelectTreeView2), new FrameworkPropertyMetadata(SelectedItemsChanged));

        private static void SelectedItemsChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
        {
            var self = (MultiSelectTreeView2)d;
            self.SelectedItemsChanged(e);
        }

        private void SelectedItemsChanged(DependencyPropertyChangedEventArgs e)
        {
            _selectedItemsHashSet = GetSelectedItems(this).OfType<object>().ToImmutableHashSet();
        }

        private bool IsSelected(object item)
        {
            return _selectedItemsHashSet.Contains(item);
        }

        private bool IsSelected(TreeViewItem treeViewItem)
        {
            var item = ItemFromContainer(treeViewItem);
            return IsSelected(item);
        }

        public static Array GetSelectedItems(MultiSelectTreeView2 element)
        {
            return (Array)element.GetValue(SelectedItemsProperty);
        }

        public static void SetSelectedItems(MultiSelectTreeView2 element, Array value)
        {
            element.SetValue(SelectedItemsProperty, value);
        }

        private object ItemFromContainer(TreeViewItem treeViewItem)
        {
            var itemContainerGenerator = FindTreeViewItem(VisualTreeHelper.GetParent(treeViewItem))?.ItemContainerGenerator ?? ItemContainerGenerator;
            var item = itemContainerGenerator.ItemFromContainer(treeViewItem);
            return item;
        }

        private static void OnTreeViewItemGotFocus(object sender, RoutedEventArgs e)
        {
            _selectTreeViewItemOnMouseUp = null;

            if (e.OriginalSource is TreeView) return;

            var treeView = FindTreeView(e.OriginalSource as DependencyObject);
            if (treeView is null)
                return;
            
            var treeViewItem = FindTreeViewItem(e.OriginalSource as DependencyObject);
            
            if (Mouse.LeftButton == MouseButtonState.Pressed && treeView.IsSelected(treeViewItem) && Keyboard.Modifiers != ModifierKeys.Control)
            {
                _selectTreeViewItemOnMouseUp = treeViewItem;
                return;
            }

            if (Mouse.LeftButton == MouseButtonState.Pressed)
                SelectItems(treeViewItem, sender as MultiSelectTreeView2);
        }

        private static void SelectItems(TreeViewItem treeViewItem, MultiSelectTreeView2 treeView)
        {
            if (treeViewItem != null && treeView != null)
            {
                if ((Keyboard.Modifiers & (ModifierKeys.Control | ModifierKeys.Shift)) == (ModifierKeys.Control | ModifierKeys.Shift))
                {
                    SelectMultipleItemsContinuously(treeView, treeViewItem, true);
                }
                else if (Keyboard.Modifiers == ModifierKeys.Control)
                {
                    SelectMultipleItemsRandomly(treeView, treeViewItem);
                }
                else if (Keyboard.Modifiers == ModifierKeys.Shift)
                {
                    SelectMultipleItemsContinuously(treeView, treeViewItem);
                }
                else
                {
                    SelectSingleItem(treeView, treeViewItem);
                }
            }
        }

        private static void OnTreeViewItemPreviewMouseDown(object sender, MouseEventArgs e)
        {
            var treeViewItem = FindTreeViewItem(e.OriginalSource as DependencyObject);

            if (treeViewItem != null && treeViewItem.IsFocused)
                OnTreeViewItemGotFocus(sender, e);
        }

        private static void OnTreeViewItemPreviewMouseUp(object sender, MouseButtonEventArgs e)
        {
            var treeViewItem = FindTreeViewItem(e.OriginalSource as DependencyObject);

            if (treeViewItem == _selectTreeViewItemOnMouseUp)
            {
                SelectItems(treeViewItem, sender as MultiSelectTreeView2);
            }
        }

        private static TreeViewItem FindTreeViewItem(DependencyObject dependencyObject)
        {
            if (!(dependencyObject is Visual || dependencyObject is Visual3D))
                return null;

            var treeViewItem = dependencyObject as TreeViewItem;
            if (treeViewItem != null)
            {
                return treeViewItem;
            }

            return FindTreeViewItem(VisualTreeHelper.GetParent(dependencyObject));
        }

        private static void SelectSingleItem(MultiSelectTreeView2 treeView, TreeViewItem treeViewItem)
        {
            // first deselect all items
            var item = treeView.ItemFromContainer(treeViewItem);
            var selectedItems = Array.CreateInstance(treeView.GetItemType() ?? item.GetType(), 1);
            selectedItems.SetValue(item, 0);
            SetSelectedItems(treeView, selectedItems);
            treeView._startItem = item;
        }

        private static MultiSelectTreeView2 FindTreeView(DependencyObject dependencyObject)
        {
            if (dependencyObject == null)
            {
                return null;
            }

            var treeView = dependencyObject as MultiSelectTreeView2;

            return treeView ?? FindTreeView(VisualTreeHelper.GetParent(dependencyObject));
        }

        private static void SelectMultipleItemsRandomly(MultiSelectTreeView2 treeView, TreeViewItem treeViewItem)
        {
            var isSelected = treeView.IsSelected(treeViewItem);
            var item = treeView.ItemFromContainer(treeViewItem);
            if (isSelected)
            {
                var selectedItems = GetSelectedItems(treeView)
                    .OfType<object>()
                    .Except(new []{item})
                    .ToArray();
                var array = Array.CreateInstance(treeView.GetItemType() ?? item.GetType(), selectedItems.Length);
                Array.Copy(selectedItems, array, selectedItems.Length);
                SetSelectedItems(treeView, array);
            }
            else
            {
                var selectedItems = GetSelectedItems(treeView)
                    .OfType<object>()
                    .Concat(new []{item})
                    .ToArray();
                var array = Array.CreateInstance(treeView.GetItemType() ?? item.GetType(), selectedItems.Length);
                Array.Copy(selectedItems, array, selectedItems.Length);
                SetSelectedItems(treeView, array);
            }

            if (treeView._startItem == null || Keyboard.Modifiers == ModifierKeys.Control)
            {
                if (isSelected is false)
                {
                    treeView._startItem = item;
                }
            }
            else
            {
                if (GetSelectedItems(treeView).Length == 0)
                {
                    treeView._startItem = null;
                }
            }
        }

        private static void SelectMultipleItemsContinuously(MultiSelectTreeView2 treeView, TreeViewItem treeViewItem, bool shiftControl = false)
        {
            var startItem = treeView._startItem;
            var treeViewItemItem = treeView.ItemFromContainer(treeViewItem);
            if (startItem != null)
            {
                if (startItem == treeViewItemItem)
                {
                    SelectSingleItem(treeView, treeViewItem);
                    return;
                }

                var items = FindTreeViewItem(VisualTreeHelper.GetParent(treeViewItem))?.Items ?? FindTreeView(treeViewItem).Items;

                if (items.Contains(startItem) is false)
                {
                    SelectSingleItem(treeView, treeViewItem);
                    return;
                }

                var selectedItems = new List<object>();
                bool isBetween = false;
                foreach (var item in items)
                {
                    if (item == treeViewItemItem || item == startItem)
                    {
                        // toggle to true if first element is found and
                        // back to false if last element is found
                        isBetween = !isBetween;

                        selectedItems.Add(item);
                        continue;
                    }

                    if (isBetween)
                    {
                        selectedItems.Add(item);
                        continue;
                    }

                    if (!shiftControl)
                        selectedItems.Remove(item);
                }

                var array = Array.CreateInstance(treeView.GetItemType() ?? treeViewItemItem.GetType(), selectedItems.Count);
                Array.Copy(selectedItems.ToArray(), array, selectedItems.Count);
                SetSelectedItems(treeView, array);
            }
        }

        private Type GetItemType()
        {
            var bindingExpression = GetBindingExpression(SelectedItemsProperty);
            if (bindingExpression is null)
                return null;
            
            return bindingExpression.ResolvedSource.GetType()
                .GetProperty(bindingExpression.ResolvedSourcePropertyName)?
                .PropertyType.GetElementType();
        }
    }
}
