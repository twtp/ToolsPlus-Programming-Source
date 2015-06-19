using System;
using System.Collections.Generic;
using System.Text;

namespace TP.Warehousing.White.Pod {
  public static class Controller {
    private static int podAddress;
    private static CarouselOrientation podType;
    private static bool podActive;
    private static int currentWarehouseID;
    private static DeviceListing deviceListing;
    private static LinkageListing linkageListing;

    private static DateTime lastButtonPushTime = DateTime.Now;

    public static event MotorFaultEventHandler MotorFault;
    public static event PhotoEyeFaultEventHandler PhotoEyeFault;
    public static event KeypadModeFaultEventHandler KeypadModeFault;
    public static event EmergencyStopEventHandler EmergencyStop;
    public static event CarouselOffEventHandler CarouselOff;
    public static event NoFaultsEventHandler NoFaults;

    public static event LightreeButtonEventHandler LightreeButton;
    public static event OtherTextMessageEventHandler OtherTextMessage;

    public static event CarouselStatusUpdateEventHandler CarouselStatusUpdate;

    public static event LightActivationEventHandler LightActivation;
    public static event LightDeactivationEventHandler LightDeactivation;

    static Controller() {
      currentWarehouseID = TP.Base.Config.Warehouse.Address;

      // sets podActive for the rest of the program lifetime
      bool ret = buildDevicesFromConfig();
      if (ret == true) {
        TP.Base.Logger.Log("Pod initializing");
        Interface.CIC332.FaultsReceived += new Interface.CIC332FaultsReceivedEventHandler(CIC332_FaultsReceived);
        Interface.CIC332.StatusReceived += new Interface.CIC332StatusReceivedEventHandler(CIC332_StatusReceived);
        Interface.CIC332.TextReceived += new Interface.CIC332TextMessageReceivedEventHandler(CIC332_TextReceived);
        Interface.CIC332.LEDIlluminated += new Interface.LEDIlluminatedEventHandler(CIC332_LEDIlluminated);
        Interface.CIC332.LEDCleared += new Interface.LEDClearedEventHandler(CIC332_LEDCleared);

        initializeAllDevices();
        Interface.CIC332.WaitForQueueCompletion();
      }
    }

    #region Carousels

    public static bool MoveToBin(string carouselAddr, int binNumber) {
      return deviceListing.GetCarousel(carouselAddr).MoveToBin(binNumber);
    }

    public static CarouselStatus CurrentCarouselStatus(string carouselAddr) {
      return deviceListing.GetCarousel(carouselAddr).CurrentStatus;
    }

    public static bool AskForStatus(string carouselAddr) {
      return deviceListing.GetCarousel(carouselAddr).AskForStatus();
    }

    public static bool AskForFaults(string carouselAddr) {
      return deviceListing.GetCarousel(carouselAddr).AskForFaults();
    }

    #endregion

    #region Lights

    public static bool PickPutQuantity(string carouselAddr, string div1, string div2, int quantity) {
      Device.Carousel c = deviceListing.GetCarousel(carouselAddr);
      List<LinkageListEntry> lights = linkageListing.GetLights(c);
      bool gotone = false;
      foreach (LinkageListEntry lle in lights) {
        Device.ILightsQuantities ilq = lle.LightDevice as Device.ILightsQuantities;
        if (ilq != null) {
          if (!ilq.LightQuantity(div1, div2, quantity, lle.Direction, carouselAddr)) {
            return false;
          }
          gotone = true;
        }
      }
      return gotone;
    }

    public static bool SortbarQuantity(string sortbarAddr, int lightNumber, int quantity) {
      return deviceListing.GetSortBar(sortbarAddr).Arrow(lightNumber, quantity);
    }

    public static bool SortbarMessage(string sortbarAddr, int lightNumber, string message) {
      return deviceListing.GetSortBar(sortbarAddr).Display(lightNumber, message, null);
    }

    public static bool ClearLightTree(string lightAddr) {
      return deviceListing.GetLightTree(lightAddr).Clear(null);
    }

    public static bool ClearLightTree(string lightAddr, int lightNumber) {
      return deviceListing.GetLightTree(lightAddr).Clear(lightNumber, null);
    }

    // should this just do any ILQ device?
    public static bool ClearCarouselLightTree(string carouselAddr) {
      bool gotone = false;
      Device.Carousel c = deviceListing.GetCarousel(carouselAddr);
      List<LinkageListEntry> lights = linkageListing.GetLights(c, DeviceType.LightTree);
      foreach (LinkageListEntry lle in lights) {
        gotone = lle.LightDevice.Clear(carouselAddr);
      }
      return gotone;
    }

    // should this just do any ILQ device?
    public static bool ClearCarouselLightTree(string carouselAddr, int lightNumber) {
      bool gotone = false;
      Device.Carousel c = deviceListing.GetCarousel(carouselAddr);
      List<LinkageListEntry> lights = linkageListing.GetLights(c, DeviceType.LightTree);
      foreach (LinkageListEntry lle in lights) {
        gotone = lle.LightDevice.Clear(lightNumber, carouselAddr);
      }
      return gotone;
    }

    public static bool ClearSortBar(string lightAddr) {
      return deviceListing.GetSortBar(lightAddr).Clear(null);
    }

    public static bool ClearSortBar(string lightAddr, int lightNumber) {
      return deviceListing.GetSortBar(lightAddr).Clear(lightNumber, null);
    }

    public static int SortbarDisplayCount(string sortbarAddr) {
      return deviceListing.GetSortBar(sortbarAddr).NumberOfDisplays;
    }

    public static bool LightTreeMessage(string carouselAddr, int lightNumber, string message) {
      Device.Carousel c = deviceListing.GetCarousel(carouselAddr);
      List<LinkageListEntry> lights = linkageListing.GetLights(c);
      foreach (LinkageListEntry lle in lights) {
        if (lle.LightDevice.DeviceType == DeviceType.LightTree) {
          return lle.LightDevice.DisplayArbitraryMessage(lightNumber, message, carouselAddr);
        }
      }
      return false;
    }

    #endregion

    #region Locking

    public static bool IsDeviceAvailable(DeviceType type, string address) {
      return !deviceListing.GetDevice(type, address).IsLocked;
    }

    public static bool IsDeviceLocked(DeviceType type, string address) {
      return deviceListing.GetDevice(type, address).IsLocked;
    }

    public static int DEBUG_DEVICE_LOCKED_BY(DeviceType type, string address) {
      return deviceListing.GetDevice(type, address).LockedByActionID;
    }

    public static bool IsDeviceLockedBy(DeviceType type, string address, int actionID) {
      return deviceListing.GetDevice(type, address).LockedByActionID == actionID;
    }

    public static bool LockDevice(DeviceType type, string address, int byActionID) {
      return deviceListing.GetDevice(type, address).LockDevice(byActionID);
    }

    public static bool UnlockDevice(DeviceType type, string address) {
      return deviceListing.GetDevice(type, address).UnlockDevice();
    }

    #endregion

    public static List<string> EnumerateDevices(DeviceType type) {
      List<Device.BaseDevice> all = deviceListing.GetAllDevices(type);
      List<string> retval = new List<string>();
      all.ForEach(delegate(Device.BaseDevice bd) { retval.Add(bd.Address); });
      return retval;
    }

    #region CIC event handlers

    private static void CIC332_StatusReceived(object sender, Interface.StatusMessageEventArgs e) {
      if (CarouselStatusUpdate != null) {
        CarouselStatusUpdateEventArgs args = new CarouselStatusUpdateEventArgs(
            e.CarouselAddress,
            e.CarouselMode,
            e.CurrentPosition,
            e.CarouselStatus
          );
        CarouselStatusUpdate(null, args);
      }
      if (e.CarouselStatus == CarouselStatus.Uninitialized) {
        // first run, we should home
        MoveToBin(e.CarouselAddress, 1);
      }
    }

    private static void CIC332_FaultsReceived(object sender, Interface.FaultsMessageEventArgs e) {
      int faultCount = 0;
      if (e.MotorError) {
        faultCount++;
        if (MotorFault != null) {
          MotorFault(null, new CarouselFaultEventArgs(e.CarouselAddress));
        }
      }

      if (e.PhotoeyeAnyError) {
        faultCount++;
        if (PhotoEyeFault != null) {
          PhotoEyeFault(null, new CarouselFaultEventArgs(e.CarouselAddress));
        }
      }

      if (e.KeypadState != KeypadState.Host) {
        faultCount++;
        if (KeypadModeFault != null) {
          KeypadModeFault(null, new CarouselFaultEventArgs(e.CarouselAddress));
        }
      }

      if (e.EmergencyStopEngaged) {
        faultCount++;
        if (EmergencyStop != null) {
          EmergencyStop(null, new CarouselFaultEventArgs(e.CarouselAddress));
        }
      }

      if (e.CarouselError) {
        faultCount++;
        if (CarouselOff != null) {
          CarouselOff(null, new CarouselFaultEventArgs(e.CarouselAddress));
        }
      }

      if (faultCount == 0) {
        if (NoFaults != null) {
          NoFaults(null, new CarouselFaultEventArgs(e.CarouselAddress));
        }
      }
    }

    private static void CIC332_TextReceived(object sender, Interface.TextMessageEventArgs e) {
      if (e.MessageText == "A DONE") {
        DateTime now = DateTime.Now;
        TimeSpan ts = now - lastButtonPushTime;
        if (ts.TotalMilliseconds < 2000) {
          // ignore duplicate input
        }
        else {
          if (LightreeButton != null) {
            LightreeButton(null, new TextMessageEventArgs(e.FromLightTree));
          }
        }
        lastButtonPushTime = now;
      }
      else {
        // TODO: should check for the "testing complete" message, then
        // re-initialize the device
        if (OtherTextMessage != null) {
          OtherTextMessage(null, new TextMessageEventArgs(e.FromLightTree));
        }
      }
    }

    private static void CIC332_LEDIlluminated(object sender, Interface.LEDIlluminatedEventArgs e) {
      // does the pod need to keep track of anything here?
      if (LightActivation != null) {
        LightActivation(null, new LightActivationEventArgs(e));
      }
    }
    private static void CIC332_LEDCleared(object sender, Interface.LEDClearedEventArgs e) {
      // does the pod need to keep track of anything here?
      //TP.Base.Logger.Log("pod reacting to LEDCleared event");
      if (LightDeactivation != null) {
        //TP.Base.Logger.Log("pod propogating event");
        LightDeactivation(null, new LightDeactivationEventArgs(e));
      }
    }

    #endregion

    private static bool buildDevicesFromConfig() {
      podActive = TP.Base.Config.Pod.IsConfigured;

      if (podActive == false) {
        return false;
      }

      podAddress = TP.Base.Config.Pod.Address;
      switch (TP.Base.Config.Pod.Orientation) {
        case TP.Base.ConfigParam.CarouselOrientation.Horizontal:
          podType = CarouselOrientation.Horizontal;
          break;
        case TP.Base.ConfigParam.CarouselOrientation.Vertical:
          podType = CarouselOrientation.Vertical;
          break;
      }

      deviceListing = new DeviceListing();
      linkageListing = new LinkageListing();

      foreach (TP.Base.ConfigParam.Carousel cfg in TP.Base.Config.Carousel) {
        deviceListing.AddDevice(new Device.Carousel(cfg.Address, cfg.Bins));
      }

      foreach (TP.Base.ConfigParam.LighTree cfg in TP.Base.Config.LighTree) {
        deviceListing.AddDevice(new Device.LightTree(cfg.Address, cfg.Lights, cfg.Chars));
      }

      foreach (TP.Base.ConfigParam.SortBar cfg in TP.Base.Config.SortBar) {
        deviceListing.AddDevice(new Device.SortBar(cfg.Address, cfg.Lights, cfg.Chars));
      }

      foreach (TP.Base.ConfigParam.CharacterStrip cfg in TP.Base.Config.CharacterStrip) {
        deviceListing.AddDevice(new Device.CharacterStrip(cfg.Address, cfg.Lights, cfg.Chars));
      }

      foreach (TP.Base.ConfigParam.SurfaceDisplay cfg in TP.Base.Config.SurfaceDisplay) {
        deviceListing.AddDevice(new Device.SurfaceDisplay(cfg.Address, cfg.Lights, cfg.Chars));
      }

      deviceListing.AddDevice(new Device.Operator());

      foreach (TP.Base.ConfigParam.PodLinkage cfg in TP.Base.Config.PodLinkage) {
        Device.Carousel c = deviceListing.GetCarousel(cfg.Carousel);
        Device.LightDevice ld = null;
        switch (cfg.LightType) {
          case TP.Base.ConfigParam.LightType.LighTree:
            ld = deviceListing.GetLightTree(cfg.LightAddr);
            break;
          case TP.Base.ConfigParam.LightType.CharacterStrip:
            ld = deviceListing.GetCharacterStrip(cfg.LightAddr);
            break;
          case TP.Base.ConfigParam.LightType.SurfaceDisplay:
            ld = deviceListing.GetSurfaceDisplay(cfg.LightAddr);
            break;
          default:
            throw new ApplicationException("oops");
        }
        linkageListing.AddLinkage(c, ld, cfg.Direction);
      }

      return true;
    }

    private static void initializeAllDevices() {
      TP.Base.Logger.Log("Pod: Beginning device initialization");
      foreach (Device.BaseDevice bd in deviceListing.GetAllDevices()) {
        if (!bd.Initialize()) {
          TP.Base.Logger.Log("Pod: ERROR INITIALIZING DEVICE " + bd.DeviceType.ToString() + " address " + bd.Address);
        }
      }
      TP.Base.Logger.Log("Pod: Device initialization completed");
    }

    private static void waitForCarouselCompletion() {
      List<Device.Carousel> carousels = deviceListing.GetAllCarousels();
      bool homed = false;
      while (!homed) {
        homed = true;
        foreach (Device.Carousel c in carousels) {
          if (c.CurrentStatus == CarouselStatus.Uninitialized) {
            homed = false;
            break;
          }
        }
        if (!homed) {
          System.Threading.Thread.Sleep(30);
        }
      }
    }

    public static int PodAddress { get { return podAddress; } }
    public static CarouselOrientation PodType { get { return podType; } }
    public static int WarehouseAddress { get { return currentWarehouseID; } }
    public static bool PodActive { get { return podActive; } }
  }
}
