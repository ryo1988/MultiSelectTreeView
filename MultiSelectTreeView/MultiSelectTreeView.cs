﻿using System;
using System.Collections;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Media3D;

namespace MultiSelectTreeView
{
    public class MultiSelectTreeView : TreeView
    {
        public MultiSelectTreeView()
        {
            PreviewMouseLeftButtonDown += OnTreeViewItemPreviewMouseDown;
            PreviewMouseLeftButtonUp += OnTreeViewItemPreviewMouseUp;
        }

        private TreeViewItem _selectTreeViewItemOnMouseUp;
        private bool _bringIntoViewWhenSelected = true;


        public static readonly DependencyProperty IsItemSelectedProperty = DependencyProperty.RegisterAttached("IsItemSelected", typeof(Boolean), typeof(MultiSelectTreeView), new PropertyMetadata(false, OnIsItemSelectedPropertyChanged));

        public static bool GetIsItemSelected(TreeViewItem element)
        {
            return (bool)element.GetValue(IsItemSelectedProperty);
        }

        public static void SetIsItemSelected(TreeViewItem element, Boolean value)
        {
            if (element == null) return;

            element.SetValue(IsItemSelectedProperty, value);
        }

        private static void OnIsItemSelectedPropertyChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
        {
            var treeViewItem = d as TreeViewItem;
            var treeView = FindTreeView(treeViewItem);
            if (treeViewItem != null && treeView != null)
            {
                var selectedItems = GetSelectedItems(treeView);
                if (GetIsItemSelected(treeViewItem))
                {
                    selectedItems?.Add(treeViewItem.Header);
                    if (treeView._bringIntoViewWhenSelected)
                    {
                        treeViewItem.BringIntoView();
                    }
                }
                else
                {
                    selectedItems?.Remove(treeViewItem.Header);
                }
            }
        }

        public static readonly DependencyProperty SelectedItemsProperty = DependencyProperty.RegisterAttached("SelectedItems", typeof(IList), typeof(MultiSelectTreeView));

        public static IList GetSelectedItems(TreeView element)
        {
            return (IList)element.GetValue(SelectedItemsProperty);
        }

        public static void SetSelectedItems(TreeView element, IList value)
        {
            element.SetValue(SelectedItemsProperty, value);
        }

        private static readonly DependencyProperty StartItemProperty = DependencyProperty.RegisterAttached("StartItem", typeof(TreeViewItem), typeof(MultiSelectTreeView));


        private static TreeViewItem GetStartItem(TreeView element)
        {
            return (TreeViewItem)element.GetValue(StartItemProperty);
        }

        private static void SetStartItem(TreeView element, TreeViewItem value)
        {
            element.SetValue(StartItemProperty, value);
        }


        private static void OnTreeViewItemGotFocus(object sender, RoutedEventArgs e)
        {
            var treeView = (MultiSelectTreeView)sender;
            
            treeView._selectTreeViewItemOnMouseUp = null;

            if (e.OriginalSource is TreeView) return;

            var treeViewItem = FindTreeViewItem(e.OriginalSource as DependencyObject);
            if (Mouse.LeftButton == MouseButtonState.Pressed && GetIsItemSelected(treeViewItem) && Keyboard.Modifiers != ModifierKeys.Control)
            {
                treeView._selectTreeViewItemOnMouseUp = treeViewItem;
                return;
            }

            if (Mouse.LeftButton == MouseButtonState.Pressed)
                SelectItems(treeViewItem, sender as MultiSelectTreeView);
        }

        private static void SelectItems(TreeViewItem treeViewItem, MultiSelectTreeView treeView)
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

            if (treeViewItem != null)
                OnTreeViewItemGotFocus(sender, e);
        }

        private static void OnTreeViewItemPreviewMouseUp(object sender, MouseButtonEventArgs e)
        {
            var treeViewItem = FindTreeViewItem(e.OriginalSource as DependencyObject);
            var treeView = (MultiSelectTreeView)sender;

            if (treeViewItem == treeView._selectTreeViewItemOnMouseUp)
            {
                SelectItems(treeViewItem, sender as MultiSelectTreeView);
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

        private static void SelectSingleItem(TreeView treeView, TreeViewItem treeViewItem)
        {
            // first deselect all items
            DeSelectAllItems(treeView, null, treeViewItem);
            SetIsItemSelected(treeViewItem, true);
            SetStartItem(treeView, treeViewItem);
        }

        private static void DeSelectAllItems(TreeView treeView, TreeViewItem treeViewItem, TreeViewItem ignoreItem)
        {
            if (treeView != null)
            {
                for (int i = 0; i < treeView.Items.Count; i++)
                {
                    var item = treeView.ItemContainerGenerator.ContainerFromIndex(i) as TreeViewItem;
                    if (item != null)
                    {
                        if (item != ignoreItem)
                            SetIsItemSelected(item, false);
                        DeSelectAllItems(null, item, ignoreItem);
                    }
                }
            }
            else
            {
                for (int i = 0; i < treeViewItem.Items.Count; i++)
                {
                    var item = treeViewItem.ItemContainerGenerator.ContainerFromIndex(i) as TreeViewItem;
                    if (item != null)
                    {
                        if (item != ignoreItem)
                            SetIsItemSelected(item, false);
                        DeSelectAllItems(null, item, ignoreItem);
                    }
                }
            }
        }

        private static MultiSelectTreeView FindTreeView(DependencyObject dependencyObject)
        {
            if (dependencyObject == null)
            {
                return null;
            }

            var treeView = dependencyObject as MultiSelectTreeView;

            return treeView ?? FindTreeView(VisualTreeHelper.GetParent(dependencyObject));
        }

        private static void SelectMultipleItemsRandomly(TreeView treeView, TreeViewItem treeViewItem)
        {
            SetIsItemSelected(treeViewItem, !GetIsItemSelected(treeViewItem));
            if (GetStartItem(treeView) == null || Keyboard.Modifiers == ModifierKeys.Control)
            {
                if (GetIsItemSelected(treeViewItem))
                {
                    SetStartItem(treeView, treeViewItem);
                }
            }
            else
            {
                if (GetSelectedItems(treeView).Count == 0)
                {
                    SetStartItem(treeView, null);
                }
            }
        }

        private static void SelectMultipleItemsContinuously(MultiSelectTreeView treeView, TreeViewItem treeViewItem, bool shiftControl = false)
        {
            TreeViewItem startItem = GetStartItem(treeView);
            if (startItem != null)
            {
                if (startItem == treeViewItem)
                {
                    SelectSingleItem(treeView, treeViewItem);
                    return;
                }

                ICollection<TreeViewItem> allItems = new List<TreeViewItem>();
                GetAllItems(treeView, null, allItems);
                //DeSelectAllItems(treeView, null);
                treeView._bringIntoViewWhenSelected = false;
                bool isBetween = false;
                foreach (var item in allItems)
                {
                    if (item == treeViewItem || item == startItem)
                    {
                        // toggle to true if first element is found and
                        // back to false if last element is found
                        isBetween = !isBetween;

                        // set boundary element
                        SetIsItemSelected(item, true);
                        continue;
                    }

                    if (isBetween)
                    {
                        SetIsItemSelected(item, true);
                        continue;
                    }

                    if (!shiftControl)
                        SetIsItemSelected(item, false);
                }
                treeView._bringIntoViewWhenSelected = true;
            }
        }

        private static void GetAllItems(TreeView treeView, TreeViewItem treeViewItem, ICollection<TreeViewItem> allItems)
        {
            if (treeView != null)
            {
                for (int i = 0; i < treeView.Items.Count; i++)
                {
                    var item = treeView.ItemContainerGenerator.ContainerFromIndex(i) as TreeViewItem;
                    if (item != null)
                    {
                        allItems.Add(item);
                        GetAllItems(null, item, allItems);
                    }
                }
            }
            else
            {
                for (int i = 0; i < treeViewItem.Items.Count; i++)
                {
                    var item = treeViewItem.ItemContainerGenerator.ContainerFromIndex(i) as TreeViewItem;
                    if (item != null)
                    {
                        allItems.Add(item);
                        GetAllItems(null, item, allItems);
                    }
                }
            }
        }
    }
}
