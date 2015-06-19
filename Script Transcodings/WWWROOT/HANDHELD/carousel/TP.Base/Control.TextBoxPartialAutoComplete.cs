using System;
using System.Collections.Generic;
using System.Text;

using System.Windows.Forms;
using System.Drawing;
using System.ComponentModel;

namespace TP.Base.Control {
  public partial class TextBoxPartialAutoComplete : TextBox {

    private ListBox list;
    private Form popup;
    private TextBoxPartialAutoComplete.WinHook hook;


    public TextBoxPartialAutoComplete() {
      // popup form to hold the list
      this.popup = new Form();
      this.popup.StartPosition = FormStartPosition.Manual;
      this.popup.ShowInTaskbar = false;
      this.popup.FormBorderStyle = FormBorderStyle.None;
      this.popup.TopMost = true;
      this.popup.Deactivate += new EventHandler(this.popup_Deactivate);

      // list box to hold matching items
      this.list = new ListBox();
      this.list.Cursor = Cursors.Hand;
      this.list.BorderStyle = BorderStyle.None;
      this.list.SelectedIndexChanged += new EventHandler(this.list_SelectedIndexChanged);
      this.list.MouseDown += new MouseEventHandler(this.list_MouseDown);
      this.list.ItemHeight = 14;
      this.list.DrawMode = DrawMode.OwnerDrawFixed;
      this.list.DrawItem += new DrawItemEventHandler(list_DrawItem);
      this.list.Dock = DockStyle.Fill;
      this.popup.Controls.Add(this.list);

      // default triggers
      this.triggers.Add(new TextLengthTrigger(0));
      this.triggers.Add(new ShortCutTrigger(Keys.Enter, TriggerState.SelectAndConsume));
      this.triggers.Add(new ShortCutTrigger(Keys.Tab, TriggerState.Select));
      this.triggers.Add(new ShortCutTrigger(Keys.Control | Keys.Space, TriggerState.ShowAndConsume));
      this.triggers.Add(new ShortCutTrigger(Keys.Escape, TriggerState.HideAndConsume));
    }

    protected virtual bool DefaultCmdKey(ref Message msg, Keys keyData) {
      bool val = base.ProcessCmdKey(ref msg, keyData);

      if (this.TriggersEnabled) {
        switch (this.Triggers.OnCommandKey(keyData)) {
          case TriggerState.ShowAndConsume:
            val = true;
            this.ShowList();
            break;
          case TriggerState.Show:
            this.ShowList();
            break;
          case TriggerState.HideAndConsume:
            val = true;
            this.HideList();
            break;
          case TriggerState.Hide:
            this.HideList();
            break;
          case TriggerState.SelectAndConsume:
            if (this.popup.Visible == true) {
              val = true;
              this.SelectCurrentItem();
            }
            break;
          case TriggerState.Select:
            if (this.popup.Visible == true) {
              this.SelectCurrentItem();
            }
            break;
          default:
            break;
        }
      }

      return val;
    }

    protected override bool ProcessCmdKey(ref Message msg, Keys keyData) {
      switch (keyData) {
        case Keys.Up:
          this.Mode = EntryMode.List;
          if (this.list.SelectedIndex > 0) {
            this.list.SelectedIndex--;
          }
          return true;
        //break;
        case Keys.Down:
          this.Mode = EntryMode.List;
          if (this.list.SelectedIndex < this.list.Items.Count - 1) {
            this.list.SelectedIndex++;
          }
          return true;
        //break;
        default:
          return DefaultCmdKey(ref msg, keyData);
        //break;
      }
    }

    protected override void OnTextChanged(EventArgs e) {
      base.OnTextChanged(e);

      if (this.TriggersEnabled) {
        switch (this.Triggers.OnTextChanged(this.Text)) {
          case TriggerState.Show:
            this.ShowList();
            break;
          case TriggerState.Hide:
            this.HideList();
            break;
          default:
            this.UpdateList();
            break;
        }
      }
    }

    protected override void OnLostFocus(EventArgs e) {
      base.OnLostFocus(e);

      if (!(this.Focused || this.popup.Focused || this.list.Focused)) {
        this.HideList();
      }
    }

    protected virtual void SelectCurrentItem() {
      if (this.list.SelectedIndex == -1) {
        return;
      }

      this.Focus();
      this.Text = this.list.SelectedItem.ToString();
      if (this.Text.Length > 0) {
        this.SelectionStart = this.Text.Length;
      }

      this.HideList();
    }

    protected virtual void ShowList() {
      if (this.popup.Visible == false) {
        this.list.SelectedIndex = -1;
        this.UpdateList();
        Point p = this.PointToScreen(new Point(0, 0));
        p.X += this.PopupOffset.X;
        p.Y += this.Height + this.PopupOffset.Y;
        this.popup.Location = p;
        if (this.list.Items.Count > 0) {
          this.popup.Show();
          if (this.hook == null) {
            this.hook = new WinHook(this);
            this.hook.AssignHandle(this.FindForm().Handle);
          }
          this.Focus();
        }
      }
      else {
        this.UpdateList();
      }
    }

    protected virtual void HideList() {
      this.Mode = EntryMode.Text;
      if (this.hook != null) {
        this.hook.ReleaseHandle();
      }
      this.hook = null;
      this.popup.Hide();
    }

    protected virtual void UpdateList() {
      object selectedItem = this.list.SelectedItem;

      this.list.Items.Clear();
      this.list.Items.AddRange(this.FilterList(this.Items).ToObjectArray());

      if (selectedItem != null && this.list.Items.Contains(selectedItem)) {
        EntryMode oldMode = this.Mode;
        this.Mode = EntryMode.List;
        this.list.SelectedItem = selectedItem;
        this.Mode = oldMode;
      }

      if (this.list.Items.Count == 0) {
        this.HideList();
      }
      else {
        int visItems = this.list.Items.Count;
        if (visItems > 8) {
          visItems = 8;
        }

        this.popup.Height = (visItems * this.list.ItemHeight) + 2;
        switch (this.BorderStyle) {
          case BorderStyle.FixedSingle:
            this.popup.Height += 2;
            break;
          case BorderStyle.Fixed3D:
            this.popup.Height += 4;
            break;
          case BorderStyle.None:
          default:
            break;
        }

        this.popup.Width = this.PopupWidth;

        if (this.list.Items.Count > 0 && this.list.SelectedIndex == -1) {
          EntryMode oldMode = this.Mode;
          this.Mode = EntryMode.List;
          this.list.SelectedIndex = 0;
          this.Mode = oldMode;
        }
      }

    }

    protected virtual AutoCompleteEntryCollection FilterList(AutoCompleteEntryCollection list) {
      AutoCompleteEntryCollection newList = new AutoCompleteEntryCollection();
      foreach (AutoCompleteEntry entry in list) {
        foreach (string match in entry.MatchStrings) {
          //if (match.ToUpper().StartsWith(this.Text.ToUpper())) {
          if (this.entryMatches(entry)) {
            newList.Add(entry);
            break;
          }
        }
      }
      return newList;
    }

    private bool entryMatches(AutoCompleteEntry entry) {
      foreach (string match in entry.MatchStrings) {
        switch (this.partialMatch) {
          case false:
            if (match.ToUpper().StartsWith(this.Text.ToUpper())) {
              return true;
            }
            break;
          case true:
            if (match.ToUpper().Contains(this.Text.ToUpper())) {
              return true;
            }
            break;
        }
      }
      return false;
    }

    #region Properties

    private TextBoxPartialAutoComplete.EntryMode mode = EntryMode.Text;
    [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
    [Browsable(false)]
    public TextBoxPartialAutoComplete.EntryMode Mode {
      get { return this.mode; }
      set { this.mode = value; }
    }

    private AutoCompleteEntryCollection items = new AutoCompleteEntryCollection();
    [System.ComponentModel.Editor(typeof(AutoCompleteEntryCollection.AutoCompleteEntryCollectionEditor), typeof(System.Drawing.Design.UITypeEditor))]
    [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
    [Browsable(true)]
    public AutoCompleteEntryCollection Items {
      get { return this.items; }
      set { this.items = value; }
    }

    private AutoCompleteTriggerCollection triggers = new AutoCompleteTriggerCollection();
    [System.ComponentModel.Editor(typeof(AutoCompleteTriggerCollection.AutoCompleteTriggerCollectionEditor), typeof(System.Drawing.Design.UITypeEditor))]
    [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
    [Browsable(false)]
    public AutoCompleteTriggerCollection Triggers {
      get { return this.triggers; }
      set { this.triggers = value; }
    }

    [Browsable(true)]
    [Description("The width of the popup (-1 will auto-size the popup to the width of the textbox).")]
    public int PopupWidth {
      get { return this.popup.Width; }
      set { this.popup.Width = value == -1 ? this.Width : value; }
    }

    public BorderStyle PopupBorderStyle {
      get { return this.list.BorderStyle; }
      set { this.list.BorderStyle = value; }
    }

    private Point popOffset = new Point(12, 0);
    [Description("The popup defaults to the lower left edge of the textbox.")]
    public Point PopupOffset {
      get { return this.popOffset; }
      set { this.popOffset = value; }
    }

    private Color popSelectBackColor = SystemColors.Highlight;
    public Color PopupSelectionBackColor {
      get { return this.popSelectBackColor; }
      set { this.popSelectBackColor = value; }
    }

    private Color popSelectForeColor = SystemColors.HighlightText;
    public Color PopupSelectionForeColor {
      get { return this.popSelectForeColor; }
      set { this.popSelectForeColor = value; }
    }

    private bool triggersEnabled = true;
    protected bool TriggersEnabled {
      get { return this.triggersEnabled; }
      set { this.triggersEnabled = value; }
    }

    [Browsable(true)]
    public override string Text {
      get { return base.Text; }
      set {
        this.TriggersEnabled = false;
        base.Text = value;
        this.TriggersEnabled = true;
      }
    }

    [Browsable(true)]
    public bool Sort {
      get { return this.list.Sorted; }
      set { this.list.Sorted = value; }
    }

    [Browsable(true)]
    private bool partialMatch = false;
    public bool PartialMatch {
      get { return this.partialMatch; }
      set { this.partialMatch = value; }
    }

    #endregion

    #region Modes

    public enum EntryMode {
      Text,
      List,
    };

    #endregion

    #region Mouse Hook

    private class WinHook : NativeWindow {
      private TextBoxPartialAutoComplete tb;

      public WinHook(TextBoxPartialAutoComplete tbox) {
        this.tb = tbox;
      }

      protected override void WndProc(ref Message m) {
        switch (m.Msg) {
          case Win32Messages.WM_LBUTTONDOWN:
          case Win32Messages.WM_LBUTTONDBLCLK:
          case Win32Messages.WM_MBUTTONDOWN:
          case Win32Messages.WM_MBUTTONDBLCLK:
          case Win32Messages.WM_RBUTTONDOWN:
          case Win32Messages.WM_RBUTTONDBLCLK:
          case Win32Messages.WM_NCLBUTTONDOWN:
          case Win32Messages.WM_NCMBUTTONDOWN:
          case Win32Messages.WM_NCRBUTTONDOWN:
            // check to see where the event took place
            {
              Form form = tb.FindForm();
              Point p = form.PointToScreen(new Point((int)m.LParam));
              Point p2 = tb.PointToScreen(new Point(0, 0));
              Rectangle rect = new Rectangle(p2, tb.Size);
              // hide the popup if it is not in the text box
              if (!rect.Contains(p)) {
                tb.HideList();
              }
            }
            break;
          case Win32Messages.WM_SIZE:
          case Win32Messages.WM_MOVE:
            tb.HideList();
            break;
          // this is the message that gets sent when a child control gets activity
          case Win32Messages.WM_PARENTNOTIFY:
            switch ((int)m.WParam) {
              case Win32Messages.WM_LBUTTONDOWN:
              case Win32Messages.WM_LBUTTONDBLCLK:
              case Win32Messages.WM_MBUTTONDOWN:
              case Win32Messages.WM_MBUTTONDBLCLK:
              case Win32Messages.WM_RBUTTONDOWN:
              case Win32Messages.WM_RBUTTONDBLCLK:
              case Win32Messages.WM_NCLBUTTONDOWN:
              case Win32Messages.WM_NCMBUTTONDOWN:
              case Win32Messages.WM_NCRBUTTONDOWN:
                // same thing as before
                {
                  Form form = tb.FindForm();
                  Point p = form.PointToScreen(new Point((int)m.LParam));
                  Point p2 = tb.PointToScreen(new Point(0, 0));
                  Rectangle rect = new Rectangle(p2, tb.Size);
                  if (!rect.Contains(p)) {
                    tb.HideList();
                  }
                }
                break;
            }
            break;
        }

        base.WndProc(ref m);
      }

      private class Win32Messages {
        public const int WM_PARENTNOTIFY = 0x210;
        public const int WM_RBUTTONDBLCLK = 0x206;
        public const int WM_RBUTTONDOWN = 0x204;
        public const int WM_RBUTTONUP = 0x205;
        public const int WM_LBUTTONDBLCLK = 0x203;
        public const int WM_LBUTTONDOWN = 0x201;
        public const int WM_LBUTTONUP = 0x202;
        public const int WM_MBUTTONDBLCLK = 0x209;
        public const int WM_MBUTTONDOWN = 0x207;
        public const int WM_MBUTTONUP = 0x208;
        public const int WM_NCLBUTTONDOWN = 0x00A1;
        public const int WM_NCLBUTTONUP = 0x00A2;
        public const int WM_NCLBUTTONDBLCLK = 0x00A3;
        public const int WM_NCRBUTTONDOWN = 0x00A4;
        public const int WM_NCRBUTTONUP = 0x00A5;
        public const int WM_NCRBUTTONDBLCLK = 0x00A6;
        public const int WM_NCMBUTTONDOWN = 0x00A7;
        public const int WM_NCMBUTTONUP = 0x00A8;
        public const int WM_NCMBUTTONDBLCLK = 0x00A9;
        public const int WM_SIZE = 0x0005;
        public const int WM_MOVE = 0x0003;
      }
    }

    #endregion

    # region Internal Events

    private void popup_Deactivate(object sender, EventArgs e) {
      if (!(this.Focused || this.popup.Focused || this.list.Focused)) {
        this.HideList();
      }
    }

    private void list_SelectedIndexChanged(object sender, EventArgs e) {
      if (this.Mode != EntryMode.List) {
        SelectCurrentItem();
      }
    }

    private void list_MouseDown(object sender, MouseEventArgs e) {
      for (int i = 0; i < this.list.Items.Count; i++) {
        if (this.list.GetItemRectangle(i).Contains(e.X, e.Y)) {
          this.list.SelectedIndex = i;
          this.SelectCurrentItem();
        }
      }
      this.HideList();
    }

    private void list_DrawItem(object sender, DrawItemEventArgs e) {
      if (e.State == DrawItemState.Selected) {
        e.Graphics.FillRectangle(new SolidBrush(this.PopupSelectionBackColor), e.Bounds);
        e.Graphics.DrawString(this.list.Items[e.Index].ToString(), e.Font, new SolidBrush(this.PopupSelectionForeColor), e.Bounds, StringFormat.GenericDefault);
      }
      else {
        e.DrawBackground();
        e.Graphics.DrawString(this.list.Items[e.Index].ToString(), e.Font, new SolidBrush(e.ForeColor), e.Bounds, StringFormat.GenericDefault);
      }
    }

    #endregion

    #region Triggers

    [Serializable]
    public enum TriggerState : int {
      None = 0,
      Show = 1,
      ShowAndConsume = 2,
      Hide = 3,
      HideAndConsume = 4,
      Select = 5,
      SelectAndConsume = 6
    };

    [Serializable]
    public abstract class AutoCompleteTrigger {
      public virtual TriggerState OnTextChanged(string text) { return TriggerState.None; }
      public virtual TriggerState OnCommandKey(Keys keyData) { return TriggerState.None; }
    }

    [Serializable]
    internal class ShortCutTrigger : AutoCompleteTrigger {
      private Keys shortCut = Keys.None;
      public Keys ShortCut {
        get { return this.shortCut; }
        set { this.shortCut = value; }
      }

      private TriggerState result = TriggerState.None;
      public TriggerState ResultState {
        get { return this.result; }
        set { this.result = value; }
      }

      public ShortCutTrigger() {
        // nothing
      }
      public ShortCutTrigger(Keys shortCutKeys, TriggerState resultState) {
        this.shortCut = shortCutKeys;
        this.result = resultState;
      }

      public override TriggerState OnCommandKey(Keys keyData) {
        if (keyData == this.ShortCut) {
          return this.ResultState;
        }
        else {
          return TriggerState.None;
        }
      }
    }

    [Serializable]
    internal class TextLengthTrigger : AutoCompleteTrigger {
      private int textLength = 2;
      public int TextLength {
        get { return this.textLength; }
        set { this.textLength = value; }
      }

      public TextLengthTrigger() {
        // nothing
      }
      public TextLengthTrigger(int length) {
        this.textLength = length;
      }

      public override TriggerState OnTextChanged(string text) {
        if (text.Length >= this.TextLength) {
          return TriggerState.Show;
        }
        else if (text.Length < this.TextLength) {
          return TriggerState.Hide;
        }
        else {
          return TriggerState.None;
        }
      }
    }

    #endregion

    #region Trigger Collections

    [Serializable]
    public class AutoCompleteTriggerCollection : System.Collections.CollectionBase {
      public AutoCompleteTrigger this[int index] {
        get { return this.InnerList[index] as AutoCompleteTrigger; }
      }

      public void Add(AutoCompleteTrigger item) {
        this.InnerList.Add(item);
      }

      public void Remove(AutoCompleteTrigger item) {
        this.InnerList.Remove(item);
      }

      public virtual TriggerState OnTextChanged(string text) {
        foreach (AutoCompleteTrigger trigger in this.InnerList) {
          TriggerState state = trigger.OnTextChanged(text);
          if (state != TriggerState.None) {
            return state;
          }
        }
        return TriggerState.None;
      }

      public virtual TriggerState OnCommandKey(Keys keyData) {
        foreach (AutoCompleteTrigger trigger in this.InnerList) {
          TriggerState state = trigger.OnCommandKey(keyData);
          if (state != TriggerState.None) {
            return state;
          }
        }
        return TriggerState.None;
      }

      public class AutoCompleteTriggerCollectionEditor : System.ComponentModel.Design.CollectionEditor {
        public AutoCompleteTriggerCollectionEditor(Type type)
          : base(type) {
          // nothing
        }
        protected override bool CanSelectMultipleInstances() { return false; }
        protected override Type[] CreateNewItemTypes() { return new Type[] { typeof(ShortCutTrigger), typeof(TextLengthTrigger) }; }
        protected override Type CreateCollectionItemType() { return typeof(ShortCutTrigger); }
      }
    }

    #endregion

    #region Entries

    public interface IAutoCompleteEntry {
      string[] MatchStrings { get; }
    }

    [Serializable]
    public class AutoCompleteEntry {
      private string[] matchStrings;
      public string[] MatchStrings {
        get {
          if (this.matchStrings == null) {
            this.matchStrings = new string[] { this.DisplayName };
          }
          return this.matchStrings;
        }
      }

      private string displayName = string.Empty;
      public string DisplayName {
        get { return this.displayName; }
        set { this.displayName = value; }
      }
      public AutoCompleteEntry() {
        // nothing
      }
      public AutoCompleteEntry(string name, params string[] matchList) {
        this.displayName = name;
        this.matchStrings = matchList;
      }
      public override string ToString() {
        return this.displayName;
      }
    }

    #endregion

    #region Entry Collections

    [Serializable]
    public class AutoCompleteEntryCollection : System.Collections.CollectionBase {
      public IAutoCompleteEntry this[int index] {
        get { return this.InnerList[index] as IAutoCompleteEntry; }
      }

      public void Add(IAutoCompleteEntry entry) {
        this.InnerList.Add(entry);
      }

      public void AddRange(System.Collections.ICollection col) {
        this.InnerList.AddRange(col);
      }

      public void Add(AutoCompleteEntry entry) {
        this.InnerList.Add(entry);
      }

      public object[] ToObjectArray() {
        return this.InnerList.ToArray();
      }

      public class AutoCompleteEntryCollectionEditor : System.ComponentModel.Design.CollectionEditor {
        public AutoCompleteEntryCollectionEditor(Type type)
          : base(type) {
          // nothing
        }
        protected override bool CanSelectMultipleInstances() { return false; }
        protected override Type CreateCollectionItemType() { return typeof(AutoCompleteEntry); }
      }
    }

    #endregion

  }
}
