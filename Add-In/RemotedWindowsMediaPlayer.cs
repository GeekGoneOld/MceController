// Written by Jonathan Dibble, Microsoft Corporation
// CODE IS PROVIDED AS-IS WITH NO WARRIENTIES EXPRESSED OR IMPLIED.

namespace VmcController.AddIn
{
    using System;
    using System.Windows.Forms;
    using System.Runtime.InteropServices;
    using WMPLib;
    using VmcController.AddIn.Commands;
    using System.Collections;


    /// <summary>
    /// This is the actual Windows Media Control.
    /// </summary>
    [System.Windows.Forms.AxHost.ClsidAttribute("{6bf52a52-394a-11d3-b153-00c04f79faa6}")]
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    public class RemotedWindowsMediaPlayer : System.Windows.Forms.AxHost, IOleServiceProvider, IOleClientSite
    {
        private WindowsMediaPlayer Player;

        /// <summary>
        /// Default Constructor.
        /// </summary>
        public RemotedWindowsMediaPlayer() : base("6bf52a52-394a-11d3-b153-00c04f79faa6")
        {
        }

        /// <summary>
        /// Used to attach the appropriate interface to Windows Media Player.
        /// In here, we call SetClientSite on the WMP Control, passing it
        /// the dotNet container (this instance.)
        /// </summary>
        protected override void AttachInterfaces()
        {
            try
            {
                Init();
                return;
            }
            catch (System.Exception ex)
            {
                System.Diagnostics.Debug.WriteLine(ex.ToString());
            }
        }

        private void Init()
        {
            //Get the IOleObject for Windows Media Player.
            IOleObject oleObject = this.GetOcx() as IOleObject;

            //Set the Client Site for the WMP control.
            oleObject.SetClientSite(this as IOleClientSite);

            Player = this.GetOcx() as WindowsMediaPlayer;
        }

        public WindowsMediaPlayer getPlayer()
        {
            if (Player != null)
            {
                Init();
            }
            return Player;
        }

        public IWMPPlaylist getNowPlaying()
        {
            if (Player != null)
            {
                if (Player.currentPlaylist == null)
                {
                    Init();
                }
                return Player.currentPlaylist;
            }
            return null;
        }

        public IWMPMedia getCurrentMediaItem()
        {
            if (Player != null)
            {
                if (Player.currentMedia == null)
                {
                    Init();
                }
                return Player.currentMedia;
            }
            return null;
        }

        public bool setNowPlaying(int index)
        {
            IWMPPlaylist playlist = getNowPlaying();
            if (playlist != null && index < playlist.count)
            {
                Player.controls.currentItem = playlist.get_Item(index);
                Player.controls.play();
                return true;
            }
            else
            {
                return false;
            }
        }    

        public WMPPlayState getPlayState()
        {
            if (Player != null)
            {
                return Player.playState;
            }
            return WMPPlayState.wmppsUndefined;
        }

        public void setShuffleMode()
        {
            if (Player != null)
            {
                Player.settings.setMode("shuffle", true);
            }
        }

        public IWMPControls getPlayerControls()
        {
            if (Player != null)
            {
                return Player.controls;
            }
            return null;
        }

        #region IOleServiceProvider Memebers - Working
        /// <summary>
        /// During SetClientSite, WMP calls this function to get the pointer to <see cref="RemoteHostInfo"/>.
        /// </summary>
        /// <param name="guidService">See MSDN for more information - we do not use this parameter.</param>
        /// <param name="riid">The Guid of the desired service to be returned.  For this application it will always match
        /// the Guid of <see cref="IWMPRemoteMediaServices"/>.</param>
        /// <returns></returns>
        IntPtr IOleServiceProvider.QueryService(ref Guid guidService, ref Guid riid)
        {
            //If we get to here, it means Media Player is requesting our IWMPRemoteMediaServices interface
            if (riid == new Guid("cbb92747-741f-44fe-ab5b-f1a48f3b2a59"))
            {
                IWMPRemoteMediaServices iwmp = (IWMPRemoteMediaServices)new RemoteHostInfo();
                return Marshal.GetComInterfaceForObject(iwmp, typeof(IWMPRemoteMediaServices));
            }
            throw new System.Runtime.InteropServices.COMException("No Interface", (int)HResults.E_NOINTERFACE);
        }
        #endregion

        #region IOleClientSite Members
        /// <summary>
        /// Not in use.  See MSDN for details.
        /// </summary>
        /// <exception cref="System.Runtime.InteropServices.COMException">E_NOTIMPL</exception>
        void IOleClientSite.SaveObject()
        {
            throw new System.Runtime.InteropServices.COMException("Not Implemented", (int)HResults.E_NOTIMPL);
        }

        /// <summary>
        /// Not in use.  See MSDN for details.
        /// </summary>
        /// <exception cref="System.Runtime.InteropServices.COMException"></exception>
        object IOleClientSite.GetMoniker(uint dwAssign, uint dwWhichMoniker)
        {
            throw new System.Runtime.InteropServices.COMException("Not Implemented", (int)HResults.E_NOTIMPL);
        }

        /// <summary>
        /// Not in use.  See MSDN for details.
        /// </summary>
        /// <exception cref="System.Runtime.InteropServices.COMException"></exception>
        object IOleClientSite.GetContainer()
        {
            return HResults.E_NOTIMPL;
        }

        /// <summary>
        /// Not in use.  See MSDN for details.
        /// </summary>
        /// <exception cref="System.Runtime.InteropServices.COMException"></exception>
        void IOleClientSite.ShowObject()
        {
            throw new System.Runtime.InteropServices.COMException("Not Implemented", (int)HResults.E_NOTIMPL);
        }

        /// <summary>
        /// Not in use.  See MSDN for details.
        /// </summary>
        /// <exception cref="System.Runtime.InteropServices.COMException"></exception>
        void IOleClientSite.OnShowWindow(bool fShow)
        {
            throw new System.Runtime.InteropServices.COMException("Not Implemented", (int)HResults.E_NOTIMPL);
        }

        /// <summary>
        /// Not in use.  See MSDN for details.
        /// </summary>
        /// <exception cref="System.Runtime.InteropServices.COMException"></exception>
        void IOleClientSite.RequestNewObjectLayout()
        {
            throw new System.Runtime.InteropServices.COMException("Not Implemented", (int)HResults.E_NOTIMPL);
        }

        #endregion
       
    }
}