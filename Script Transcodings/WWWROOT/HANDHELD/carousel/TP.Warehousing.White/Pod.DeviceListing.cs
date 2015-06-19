using System;
using System.Collections.Generic;
using System.Text;

namespace TP.Warehousing.White.Pod {
  internal class DeviceListing {

    private List<Device.BaseDevice> devices;

    public DeviceListing() {
      this.devices = new List<Device.BaseDevice>();
    }

    public void AddDevice(Device.BaseDevice newDevice) {
      this.devices.Add(newDevice);
    }

    public Device.BaseDevice GetDevice(DeviceType deviceType, string address) {
      address = Interface.CIC332.FormatAddress(address);
      foreach (Device.BaseDevice bd in this.devices) {
        if (bd.DeviceType == deviceType && bd.Address == address) {
          return bd;
        }
      }
      throw new ArgumentException("no " + deviceType.ToString() + " with address " + address + " defined!");
    }

    public List<Device.BaseDevice> GetAllDevices() {
      List<Device.BaseDevice> retval = new List<Device.BaseDevice>();
      foreach (Device.BaseDevice bd in this.devices) {
        retval.Add(bd);
      }
      return retval;
    }

    public List<Device.BaseDevice> GetAllDevices(DeviceType deviceType) {
      List<Device.BaseDevice> retval = new List<Device.BaseDevice>();
      foreach (Device.BaseDevice bd in this.devices) {
        if (bd.DeviceType == deviceType) {
          retval.Add(bd);
        }
      }
      return retval;
    }

    public Device.Carousel GetCarousel(int address) {
      return this.GetCarousel(address.ToString());
    }
    public Device.Carousel GetCarousel(string address) {
      return (Device.Carousel)this.GetDevice(DeviceType.Carousel, address);
    }
    public List<Device.Carousel> GetAllCarousels() {
      return GetAllDevices(DeviceType.Carousel).ConvertAll(new Converter<Device.BaseDevice, Device.Carousel>(delegate(Device.BaseDevice bd) { return (Device.Carousel)bd; }));
    }

    public Device.LightTree GetLightTree(int address) {
      return this.GetLightTree(address.ToString());
    }
    public Device.LightTree GetLightTree(string address) {
      return (Device.LightTree)this.GetDevice(DeviceType.LightTree, address);
    }
    public List<Device.LightTree> GetAllLightTrees() {
      return GetAllDevices(DeviceType.LightTree).ConvertAll(new Converter<Device.BaseDevice, Device.LightTree>(delegate(Device.BaseDevice bd) { return (Device.LightTree)bd; }));
    }

    public Device.SortBar GetSortBar(int address) {
      return this.GetSortBar(address.ToString());
    }
    public Device.SortBar GetSortBar(string address) {
      return (Device.SortBar)this.GetDevice(DeviceType.SortBar, address);
    }
    public List<Device.SortBar> GetAllSortBars() {
      return GetAllDevices(DeviceType.SortBar).ConvertAll(new Converter<Device.BaseDevice, Device.SortBar>(delegate(Device.BaseDevice bd) { return (Device.SortBar)bd; }));
    }

    public Device.CharacterStrip GetCharacterStrip(int address) {
      return this.GetCharacterStrip(address.ToString());
    }
    public Device.CharacterStrip GetCharacterStrip(string address) {
      return (Device.CharacterStrip)this.GetDevice(DeviceType.CharacterStrip, address); ;
    }
    public List<Device.CharacterStrip> GetAllCharacterStrips() {
      return GetAllDevices(DeviceType.CharacterStrip).ConvertAll(new Converter<Device.BaseDevice, Device.CharacterStrip>(delegate(Device.BaseDevice bd) { return (Device.CharacterStrip)bd; }));
    }

    public Device.SurfaceDisplay GetSurfaceDisplay(int address) {
      return this.GetSurfaceDisplay(address.ToString());
    }
    public Device.SurfaceDisplay GetSurfaceDisplay(string address) {
      return (Device.SurfaceDisplay)this.GetDevice(DeviceType.SurfaceDisplay, address);
    }
    public List<Device.SurfaceDisplay> GetAllSurfaceDisplays() {
      return GetAllDevices(DeviceType.SurfaceDisplay).ConvertAll(new Converter<Device.BaseDevice, Device.SurfaceDisplay>(delegate(Device.BaseDevice bd) { return (Device.SurfaceDisplay)bd; }));
    }

    public Device.Operator GetOperator(string address) {
      return (Device.Operator)this.GetDevice(DeviceType.Operator, address);
    }
    public List<Device.Operator> GetAllOperators() {
      return GetAllDevices(DeviceType.Operator).ConvertAll(new Converter<Device.BaseDevice, Device.Operator>(delegate(Device.BaseDevice bd) { return (Device.Operator)bd; }));
    }
  }
}
