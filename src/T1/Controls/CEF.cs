
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;

namespace T1.Controls
{
    public class CEF : UserControl
    {
        /*
        #region defined constants

        protected const uint WS_OVERLAPPED = 0;

        protected const uint WS_POPUP = 2147483648;

        protected const uint WS_CHILD = 1073741824;

        protected const uint WS_MINIMIZE = 536870912;

        protected const uint WS_VISIBLE = 268435456;

        protected const uint WS_DISABLED = 134217728;

        protected const uint WS_CLIPSIBLINGS = 67108864;

        protected const uint WS_CLIPCHILDREN = 33554432;

        protected const uint WS_MAXIMIZE = 16777216;

        protected const uint WS_CAPTION = 12582912;

        protected const uint WS_BORDER = 8388608;

        protected const uint WS_DLGFRAME = 4194304;

        protected const uint WS_VSCROLL = 2097152;

        protected const uint WS_HSCROLL = 1048576;

        protected const uint WS_SYSMENU = 524288;

        protected const uint WS_THICKFRAME = 262144;

        protected const uint WS_GROUP = 131072;

        protected const uint WS_TABSTOP = 65536;

        protected const uint WS_MINIMIZEBOX = 131072;

        protected const uint WS_MAXIMIZEBOX = 65536;

        protected const int GWL_STYLE = -16;

        [DllImport("user32.dll", CharSet = CharSet.None, ExactSpelling = false)]
        protected static extern bool GetClientRect(IntPtr hWnd, out CEF.RECT lpRect);

        [DllImport("user32.dll", CharSet = CharSet.None, ExactSpelling = false)]
        protected static extern uint GetWindowLong(IntPtr hWnd, int nIndex);

        [DllImport("user32.dll", CharSet = CharSet.None, ExactSpelling = false)]
        protected static extern bool GetWindowRect(IntPtr hWnd, out CEF.RECT lpRect);

        [DllImport("user32.dll", CharSet = CharSet.None, ExactSpelling = false)]
        protected static extern IntPtr SetParent(IntPtr hWndChild, IntPtr hWndNewParent);

        [DllImport("user32.dll", CharSet = CharSet.None, ExactSpelling = false)]
        protected static extern uint SetWindowLong(IntPtr hWnd, int nIndex, IntPtr dwNewLong);


        #endregion

        #region defined structures
        public struct RECT
        {
            public int Left;

            public int Top;

            public int Right;

            public int Bottom;

            public int Height
            {
                get
                {
                    return this.Bottom - this.Top;
                }
            }

            public int Width
            {
                get
                {
                    return this.Right - this.Left;
                }
            }

            public RECT(int left_, int top_, int right_, int bottom_)
            {
                this.Left = left_;
                this.Top = top_;
                this.Right = right_;
                this.Bottom = bottom_;
            }
        }
        #endregion defined structures

        #region defined enums
        public enum B1FormModes
        {
            
            fm_FIND_MODE = 0,
            fm_OK_MODE = 1,
            fm_UPDATE_MODE = 2,
            fm_ADD_MODE = 3,
            fm_VIEW_MODE = 4,
            fm_PRINT_MODE = 5,
            fm_ALL_MODE = 100
        }
        #endregion



        protected static Dictionary<string, CEF> cefList;

        

        static CEF()
        {
            CEF.cefList = new Dictionary<string, CEF>();
        }

        private CEF()
        {
        }

        protected CEF(string UserFormType)
        {
            Control.CheckForIllegalCrossThreadCalls = false;
            this._autoResetEvt = new AutoResetEvent(false);
            //this._autoFormPositionSupport = autoFormPositionSupport;
            this._B1UserFormType = UserFormType;
            this._B1FormUID = Guid.NewGuid().ToString().Substring(0,20);
            //this.BackColor = LayoutFactory.getDefaultBgColor();
        }





        protected B1FormModes _formMode = B1FormModes.fm_OK_MODE;
        public virtual B1FormModes FormMode
        {
            get
            {
                return this._formMode;
            }
            set
            {
                this._formMode = value;
                if (this.FormModeChanged != null)
                {
                    this.FormModeChanged(this, new EventArgs());
                }
            }
        }


        protected SAPbouiCOM.Form _B1Form;
        public virtual SAPbouiCOM.Form CEFForm
        {
            get
            {
                return this._B1Form;
            }
        }


        protected string _B1UserFormType;

        protected string _B1FormUID;

        protected IntPtr _B1FormPtr;

        //protected bool _autoFormPositionSupport;

        protected Thread _thread;

        protected AutoResetEvent _autoResetEvt;

        protected ApplicationContext _applicationContext;

        protected int _nonClientAreaWidthAdjust;

        protected int _nonClientAreaHeightAdjust;

        protected bool _isLoading;
        public virtual bool IsLoading
        {
            get
            {
                return this._isLoading;
            }
        }

        

        protected virtual void Activate(SAPbouiCOM.ItemEvent pVal)
        {
            try
            {
                if (!base.InvokeRequired)
                {
                    this.Refresh();
                }
                else
                {
                    CEF CEFForm = this;
                    CEF.InvokeFormActivateEventHandler invokeFormActivateEventHandler = new CEF.InvokeFormActivateEventHandler(CEFForm.Activate);
                    object[] objArray = new object[] { pVal };
                    base.BeginInvoke(invokeFormActivateEventHandler, objArray);
                }
            }
            catch (Exception exception)
            {
                ////TODO: AddException Handler
            }
        }

        public virtual void Close()
        {
            if (this._B1Form != null)
            {
                this._B1Form.Close();
            }
        }

        protected virtual Form CreateSAPForm()
        {
            this._B1Form = Form.CreateNewForm(this._B1UserFormType, this._B1FormUID);
            this._B1Form.Width = 100;
            this._B1Form.Height = 100;
            this._B1Form.Visible = false;
            return this._B1Form;
        }

                
        protected virtual Rectangle GetInitialDimensions()
        {
            Desktop desktop = B1Connector.GetB1Connector().Application.get_Desktop();
            return new Rectangle(desktop.get_Width() / 2 - base.Width / 2, desktop.get_Height() / 2 - base.Height / 2, base.Width + (this.sapForm.Width - this.sapForm.ClientWidth), base.Height + (this.sapForm.Height - this.sapForm.ClientHeight));
        }

        

        protected virtual void InitComponents()
        {
        }

        protected virtual void LoadFormSettings()
        {
            Rectangle initialDimensions = this.GetInitialDimensions();
            if (!this._autoFormPositionSupport)
            {
                this.sapForm.Width = initialDimensions.Width;
                this.sapForm.Height = initialDimensions.Height;
                this.sapForm.Left = initialDimensions.Left;
                this.sapForm.Top = initialDimensions.Top;
                return;
            }
            IDAO dAO = (new FormSettings()).DAO;
            object[] userId = new object[] { "WHERE [U_FormType]='", this.sapFormType, "' AND [U_UserId]=", B1Connector.GetB1Connector().UserId, " AND [U_DeletedF]<>'Y'" };
            List<IDTO> byWhereClause = dAO.GetByWhereClause(string.Concat(userId));
            initialDimensions = (byWhereClause.Count <= 0 ? this.GetInitialDimensions() : ((FormSettings)byWhereClause[0]).Dimension);
            this.sapForm.Width = initialDimensions.Width;
            this.sapForm.Height = initialDimensions.Height;
            this.sapForm.Left = initialDimensions.Left;
            this.sapForm.Top = initialDimensions.Top;
            if (this.sapForm.Left < 0)
            {
                this.sapForm.Left = 0;
            }
            if (this.sapForm.Top < 0)
            {
                this.sapForm.Top = 0;
            }
            if (this.sapForm.Left > B1Connector.GetB1Connector().Application.get_Desktop().get_Width())
            {
                this.sapForm.Left = 0;
            }
            if (this.sapForm.Top > B1Connector.GetB1Connector().Application.get_Desktop().get_Height())
            {
                this.sapForm.Top = 0;
            }
        }

        public virtual void LoadToSap()
        {
            this.LoadToSap(true, true);
        }

        public virtual void LoadToSap(bool wait, bool visible)
        {
            
            this._autoResetEvt.Reset();
            this._isLoading = true;
            this._thread = new Thread(() => {
                try
                {
                    this._B1Form = this.CreateSAPForm();
                    this.sapForm.Value = this.sapForm.UniqueID;
                    Dictionary<IntPtr, string> childHandles = SAPMdiHandles.MdiWindow.GetChildHandles();
                    this.sapForm.Load();
                    Dictionary<IntPtr, string> intPtrs = SAPMdiHandles.MdiWindow.GetChildHandles();
                    this.LoadFormSettings();
                    this.sapForm.Value = this.text;
                    this.sapForm.Visible = visible;
                    CEF.dotNetFormsList.Add(this.sapFormUniqueId, this);
                    this.sapFormPtr = IntPtr.Zero;
                    foreach (KeyValuePair<IntPtr, string> childHandle in intPtrs)
                    {
                        if (childHandles.ContainsKey(childHandle.Key) || !childHandle.Value.Contains(this.sapForm.UniqueID))
                        {
                            continue;
                        }
                        this.sapFormPtr = childHandle.Key;
                        break;
                    }
                    if (this.sapFormPtr == IntPtr.Zero)
                    {
                        throw new Exception("Could not find new created SAP form");
                    }
                    CEF.RECT rECT = new CEF.RECT();
                    CEF.RECT rECT1 = new CEF.RECT();
                    if (CEF.GetWindowRect(this.sapFormPtr, out rECT) && CEF.GetClientRect(this.sapFormPtr, out rECT1))
                    {
                        this._nonClientAreaWidthAdjust = rECT.Width - rECT1.Width;
                        this._nonClientAreaHeightAdjust = rECT.Height - rECT1.Height;
                    }
                    CEF.SetParent(this.Handle, this.sapFormPtr);
                    this.ChangeStyle(this.sapFormPtr);
                    this.Dock = DockStyle.Fill;
                    this.InitComponents();
                    Form u003cu003e4_this = this.sapForm;
                    CEF sAPDotNetForm = this;
                    u003cu003e4_this.AddHandler_Activate(ModeComponent.FormModes.ALL, null, new FormActiveEventHandler(sAPDotNetForm.Activate));
                    Form form = this.sapForm;
                    CEF u003cu003e4_this1 = this;
                    form.AddHandler_Resize(ModeComponent.FormModes.ALL, null, new FormResizeEventHandler(u003cu003e4_this1.Resized));
                    Form form1 = this.sapForm;
                    CEF sAPDotNetForm1 = this;
                    FormCloseEventHandler formCloseEventHandler = new FormCloseEventHandler(sAPDotNetForm1.FormClosing);
                    CEF u003cu003e4_this2 = this;
                    form1.AddHandler_Close(ModeComponent.FormModes.ALL, formCloseEventHandler, new FormCloseEventHandler(u003cu003e4_this2.FormClosed));
                    this.Resized(null);
                    this.OnBeforeShow();
                    base.Show();
                    this.OnAfterShow();
                    this._applicationContext = new ApplicationContext();
                    this._isLoading = false;
                    this._autoResetEvt.Set();
                    Application.Run(this._applicationContext);
                }
                catch (Exception exception)
                {
                    Debug.WriteMessage(string.Concat(".Net form thread crashed. Message: ", exception), Debug.DebugLevel.Exception);
                    this._isLoading = false;
                    try
                    {
                        this._autoResetEvt.Set();
                    }
                    catch
                    {
                    }
                }
            });
            this._thread.SetApartmentState(ApartmentState.STA);
            this._thread.Start();
            if (wait)
            {
                this._autoResetEvt.WaitOne();
            }
        }

        

               

        protected virtual void Resized(SAPbouiCOM.ItemEvent pVal)
        {
            try
            {
                if (!base.InvokeRequired)
                {
                    CEF.RECT rECT = new CEF.RECT();
                    CEF.RECT rECT1 = new CEF.RECT();
                    if (CEF.GetWindowRect(this.sapFormPtr, out rECT) && CEF.GetClientRect(this.sapFormPtr, out rECT1) && rECT1.Width > 0 && rECT1.Height > 0)
                    {
                        this.CEFForm.ClientWidth = (rECT.Width - rECT1.Width > this._nonClientAreaWidthAdjust ? rECT1.Width + (rECT.Width - rECT1.Width - this._nonClientAreaWidthAdjust) : rECT1.Width);
                        this.CEFForm.ClientHeight = (rECT.Height - rECT1.Height > this._nonClientAreaHeightAdjust ? rECT1.Height + (rECT.Height - rECT1.Height - this._nonClientAreaHeightAdjust) : rECT1.Height);
                        base.Size = new Size(this.CEFForm.ClientWidth, this.CEFForm.ClientHeight);
                    }
                }
                else
                {
                    CEF sAPDotNetForm = this;
                    CEF.InvokeResizedEventHandler invokeResizedEventHandler = new CEF.InvokeResizedEventHandler(sAPDotNetForm.Resized);
                    object[] objArray = new object[] { pVal };
                    base.BeginInvoke(invokeResizedEventHandler, objArray);
                }
            }
            catch (Exception exception)
            {
                Debug.WriteMessage(exception, Debug.DebugLevel.Exception);
            }
        }


        public void WaitLoaded()
        {
            while (this._isLoading)
            {
                this._autoResetEvt.WaitOne(200, true);
            }
        }

        public event EventHandler FormModeChanged;

        protected delegate void InvokeFormActivateEventHandler(SAPbouiCOM.ItemEvent pVal);

        protected delegate void InvokeResizedEventHandler(SAPbouiCOM.ItemEvent pVal);

        protected delegate void UpdateFormDelegate();


    */
    }
    
}
