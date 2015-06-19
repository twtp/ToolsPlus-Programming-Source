using System;
using System.Collections.Generic;
using System.Text;

namespace TP.Base {
  public static class Config {

    private static ConfigParam.ConfigRoot root = new ConfigParam.ConfigRoot();

    static Config() {
      // static initializer, pre-create all branches
      // this way, the program doesn't need to do anything and
      // the config will be init'ed to defaults on first read,
      // whether we call .Read() or not.
      root.SetDefaults();
    }

    // ADDING A NEW CONFIG LINE TYPE
    //  - create the class for the line parser (or list+class for multi)
    //  - add a getter here
    //  - that's it!

    public static ConfigParam.Warehouse Warehouse { get { return (ConfigParam.Warehouse)root.Branches["warehouse"]; } }
    public static ConfigParam.Database Database { get { return (ConfigParam.Database)root.Branches["database"]; } }
    public static ConfigParam.Autologin Autologin { get { return (ConfigParam.Autologin)root.Branches["autologin"]; } }
    public static ConfigParam.Scanner Scanner { get { return (ConfigParam.Scanner)root.Branches["scanner"]; } }
    public static ConfigParam.Scale Scale { get { return (ConfigParam.Scale)root.Branches["scale"]; } }
    public static ConfigParam.DistanceMeterList DistanceMeter { get { return (ConfigParam.DistanceMeterList)root.Branches["distancemeter"]; } }
    public static ConfigParam.Camera Camera { get { return (ConfigParam.Camera)root.Branches["camera"]; } }
    public static ConfigParam.CubiScan CubiScan { get { return (ConfigParam.CubiScan)root.Branches["cubiscan"]; } }
    public static ConfigParam.HeatSealer HeatSealer { get { return (ConfigParam.HeatSealer)root.Branches["heatsealer"]; } }
    public static ConfigParam.MAS200 MAS200 { get { return (ConfigParam.MAS200)root.Branches["mas200"]; } }
    public static ConfigParam.Launcher Launcher { get { return (ConfigParam.Launcher)root.Branches["launcher"]; } }
    public static ConfigParam.Locations Locations { get { return (ConfigParam.Locations)root.Branches["locations"]; } }
    public static ConfigParam.PickOptions PickOptions { get { return (ConfigParam.PickOptions)root.Branches["pickopts"]; } }
    public static ConfigParam.IdleLock IdleLock { get { return (ConfigParam.IdleLock)root.Branches["idlelock"]; } }
    public static ConfigParam.PrinterList Printer { get { return (ConfigParam.PrinterList)root.Branches["printer"]; } }
    public static ConfigParam.TextToSpeech TextToSpeech { get { return (ConfigParam.TextToSpeech)root.Branches["speech"]; } }
    public static ConfigParam.Pod Pod { get { return (ConfigParam.Pod)root.Branches["pod"]; } }
    public static ConfigParam.CarouselList Carousel { get { return (ConfigParam.CarouselList)root.Branches["carousel"]; } }
    public static ConfigParam.LighTreeList LighTree { get { return (ConfigParam.LighTreeList)root.Branches["lightree"]; } }
    public static ConfigParam.SortBarList SortBar { get { return (ConfigParam.SortBarList)root.Branches["sortbar"]; } }
    public static ConfigParam.CharacterStripList CharacterStrip { get { return (ConfigParam.CharacterStripList)root.Branches["characterstrip"]; } }
    public static ConfigParam.SurfaceDisplayList SurfaceDisplay { get { return (ConfigParam.SurfaceDisplayList)root.Branches["surfacedisplay"]; } }
    public static ConfigParam.PodLinkageList PodLinkage { get { return (ConfigParam.PodLinkageList)root.Branches["linkage"]; } }

    public static bool Read(string exePath) {
      string clientName = Interop.RDP.GetRDPClientName().ToLower();
      string rootDir = System.IO.Path.GetPathRoot(Environment.SystemDirectory);
      string exeDir = System.IO.Path.GetDirectoryName(exePath);

      List<string> filesToCheck = new List<String>();
      if (clientName != string.Empty) {
        Logger.Log("running under RDP, will check clientname configs");
        filesToCheck.Add(pathCombine(exeDir, clientName + ".conf"));
        filesToCheck.Add(pathCombine(rootDir, "WhseApps", "Configs", clientName + ".conf"));
        filesToCheck.Add(pathCombine(rootDir, clientName + ".conf"));
      }
      filesToCheck.Add(pathCombine(exeDir, "warehouse.conf"));
      filesToCheck.Add(pathCombine(rootDir, "WhseApps", "Configs", "warehouse.conf"));
      filesToCheck.Add(pathCombine(rootDir, "warehouse.conf"));
      
      foreach (string testPath in filesToCheck) {
        Logger.Log("checking config '" + testPath + "'");
        if (System.IO.File.Exists(testPath)) {
          Logger.Log("config found, running parse");
          // no try/catch here, just error?
          return parse(testPath);
        }
      }

      return false;
    }

    private static string pathCombine(params string[] paths) {
      string temp;
      temp = paths[0];
      for (int i = 1; i < paths.Length; i++) {
        temp = System.IO.Path.Combine(temp, paths[i]);
      }
      return temp;
    }

    private static bool parse(string fname) {
      string[] lines = System.IO.File.ReadAllLines(fname);
      for (int i = 0; i < lines.Length; i++) {
        string ln = lines[i].ToLower();
        ln = System.Text.RegularExpressions.Regex.Replace(ln, @"#.*$", "");
        ln = System.Text.RegularExpressions.Regex.Replace(ln, @"^\s+", "");
        ln = System.Text.RegularExpressions.Regex.Replace(ln, @"\s$", "");
        if (ln != string.Empty) {
          string[] tokens = System.Text.RegularExpressions.Regex.Split(ln, @"\s+");

          try {
            root.Parse(tokens);
          }
          catch (Exception ex) {
            throw new ArgumentException(ex.Message + " on config line " + (i+1).ToString());
          }

        }
      }

      return true;
    }

    // CONVENIENCE FUNCTIONS
    private static string machineName;
    private static string rdpClientName;
    private static string rdpClientNameOriginal;

    public static string GetMachineName() {
      return GetMachineName(true);
    }
    public static string GetMachineName(bool useCache) {
      if (machineName == null || useCache == false) {
        clientNameInit();
      }
      return machineName;
    }

    public static string GetRDPClientName() {
      return GetRDPClientName(true);
    }
    public static string GetRDPClientName(bool useCache) {
      if (machineName == null || useCache == false) {
        clientNameInit();
      }
      return rdpClientName;
    }
    public static bool RDPClientHasChanged() {
      return rdpClientNameOriginal != GetRDPClientName(false);
    }
    public static string GetRDPClientOriginalName() {
      if (machineName == null) {
        clientNameInit();
      }
      return rdpClientNameOriginal;
    }

    private static void clientNameInit() {
      machineName = Environment.MachineName.ToUpper();
      rdpClientName = Interop.RDP.GetRDPClientName();
      if (rdpClientNameOriginal == null) {
        rdpClientNameOriginal = rdpClientName;
      }
      TP.Base.Logger.Log("machine name is " + machineName + ", rdp client is " + rdpClientName);
    }

  }

  namespace ConfigParam {

    internal class RootConfigTokenAttribute : Attributes.GenericClassAttribute {
      private string rootConfigToken;
      public RootConfigTokenAttribute(string rootConfigToken) {
        this.rootConfigToken = rootConfigToken;
      }
      public string RootConfigToken { get { return this.rootConfigToken; } }
      public override object TheAttribute { get { return this.rootConfigToken; } }
    }

    internal class RootConfigAllowMultipleAttribute : Attributes.GenericClassAttribute {
      private bool rootConfigAllowMultiple;
      public RootConfigAllowMultipleAttribute(bool rootConfigAllowMultiple) {
        this.rootConfigAllowMultiple = rootConfigAllowMultiple;
      }
      public bool RootConfigAllowMultiple { get { return this.rootConfigAllowMultiple; } }
      public override object TheAttribute { get { return this.rootConfigAllowMultiple; } }
    }

    internal class ParameterConfigTokenAttribute : Attributes.GenericEnumAttribute {
      private string parameterConfigToken;
      public ParameterConfigTokenAttribute(string parameterConfigToken) {
        this.parameterConfigToken = parameterConfigToken;
      }
      public string ParameterConfigToken { get { return this.parameterConfigToken; } }
      public override object TheAttribute { get { return this.parameterConfigToken; } }
    }

    internal class ConfigRoot {
      private Dictionary<string, ConfigBranch> branches;

      public ConfigRoot() {
        this.branches = new Dictionary<string, ConfigBranch>();
      }

      public Dictionary<string, ConfigBranch> Branches { get { return this.branches; } }

      private static System.Reflection.Assembly currentAssembly = System.Reflection.Assembly.GetExecutingAssembly();

      internal bool Parse(string[] tokens) {
        if (tokens.Length % 2 == 0) {
          throw new ArgumentException("incorrect number of pairs in list");
        }

        string rootToken = tokens[0];

        // lookup class via token, instantiate and use parser to generate
        // configuration object.
        List<Type> foundClass = RootConfigTokenAttribute.GetClassesWith<RootConfigTokenAttribute>(rootToken, currentAssembly);
        if (foundClass == null || foundClass.Count != 1) {
          throw new ArgumentException("unhandled type '" + rootToken + "'");
        }
        if (!this.branches.ContainsKey(rootToken)) {
          this.branches.Add(rootToken, (ConfigBranch)Activator.CreateInstance(foundClass[0]));
        }
        return this.branches[rootToken].Parse(tokens);
      }

      internal bool SetDefaults() {
        List<Type> allBranches = RootConfigTokenAttribute.GetClassesWith<RootConfigTokenAttribute>(System.Reflection.Assembly.GetExecutingAssembly());
        foreach (Type t in allBranches) {
          string token = (string)RootConfigTokenAttribute.GetAttributeValue<RootConfigTokenAttribute>(t);
          if (!branches.ContainsKey(token)) {
            this.branches.Add(token, (ConfigBranch)Activator.CreateInstance(t));
          }
        }
        return true;
      }

    }

    public abstract class ConfigBranch {
      internal bool Parse(string[] tokens) {
        for (int i = 2; i < tokens.Length; i += 2) {
          string token = null;
          string value = null;
          try {
            token = tokens[i - 1];
            value = tokens[i];
            if (this.handleToken(token, value) == false) {
              throw new ArgumentException("invalid token '" + token + "'");
            }
          }
          catch (FormatException) {
            throw new ArgumentException("invalid value for token '" + token + "'");
          }
        }

        this.Lock = true;

        return true;
      }

      protected abstract bool handleToken(string token, string value);

      protected bool locked = false;
      internal virtual bool Lock {
        get { return this.locked; }
        set { this.locked = value; }
      }
    }

    public abstract class ConfigBranchMulti<T> : ConfigBranch, IEnumerable<T>, System.Collections.IEnumerable where T : class {
      protected List<T> _list = new List<T>();
      protected T wip = null;

      internal override bool Lock {
        get {
          return base.Lock;
        }
        set {
          if (value) {
            this._list.Add(this.wip);
            this.wip = null;
            base.Lock = false;
          }
          else {
            base.Lock = value;
          }
        }
      }

      public T this[int index] {
        get { return this._list[index]; }
      }

      public IEnumerator<T> GetEnumerator() {
        return this._list.GetEnumerator();
      }
      System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator() {
        return this._list.GetEnumerator();
      }

      public int Count { get { return this._list.Count; } }
    }

    #region ROOT -> WAREHOUSE

    [RootConfigToken("warehouse")]
    [RootConfigAllowMultiple(false)]
    public class Warehouse : ConfigBranch {
      private int address = 5;
      public int Address {
        get { return this.address; }
        internal set { this.address = value; }
      }

      protected override bool handleToken(string token, string value) {
        switch (token) {
          case "address":
            this.Address = Convert.ToInt32(value);
            return true;
          default:
            return false;
        }
      }
    }

    #endregion

    #region ROOT -> DATABASE

    [RootConfigToken("database")]
    [RootConfigAllowMultiple(false)]
    public class Database : ConfigBranch {
      private string server = "toolsplus06";
      public string Server {
        get { return this.server; }
        internal set { this.server = value; }
      }
      private string catalog = "toolsplus";
      public string Catalog {
        get { return this.catalog; }
        internal set { this.catalog = value; }
      }
      private DatabaseConnectionMethod method = DatabaseConnectionMethod.SSPI;
      public DatabaseConnectionMethod Method {
        get { return this.method; }
        internal set { this.method = value; }
      }
      private string username = null;
      public string Username {
        get { return this.username; }
        internal set { this.username = value; }
      }
      private string password = null;
      public string Password {
        get { return this.password; }
        internal set { this.password = value; }
      }
      private bool logsql = false;
      public bool LogSQL {
        get { return this.logsql; }
        internal set { this.logsql = value; }
      }

      protected override bool handleToken(string token, string value) {
        switch (token) {
          case "server":
            this.Server = value;
            return true;
          case "catalog":
            this.Catalog = value;
            return true;
          case "method":
            this.Method = (DatabaseConnectionMethod)Attributes.GenericEnumAttribute.GetEnumWith<DatabaseConnectionMethod, ParameterConfigTokenAttribute>(value);
            return true;
          case "username":
            this.Username = value;
            return true;
          case "password":
            this.Password = value;
            return true;
          case "logsql":
            this.LogSQL = Convert.ToBoolean(value);
            return true;
          default:
            return false;
        }
      }

      public string ConnectionString() {
        switch (this.Method) {
          case DatabaseConnectionMethod.SSPI:
            return string.Format("Data Source={0};Initial Catalog={1};Integrated Security=SSPI", this.Server, this.Catalog);
          case DatabaseConnectionMethod.UserPass:
            return string.Format("Data Source={0};Initial Catalog={1};User Id={2};Password={3}", this.Server, this.Catalog, this.Username, this.Password);
          default:
            throw new ApplicationException("oops");

        }
      }
    }

    #endregion

    #region ROOT -> AUTOLOGIN

    [RootConfigToken("autologin")]
    [RootConfigAllowMultiple(false)]
    public class Autologin : ConfigBranch {
      private string username = null;
      public string Username {
        get { return this.username; }
        internal set { this.username = value; }
      }
      private bool bypass = false;
      public bool Bypass {
        get { return this.bypass; }
        internal set { this.bypass = value; }
      }

      protected override bool handleToken(string token, string value) {
        switch (token) {
          case "username":
            this.Username = value;
            return true;
          case "bypass":
            this.Bypass = Convert.ToBoolean(value);
            return true;
          default:
            return false;
        }
      }
    }

    #endregion

    #region ROOT -> SCANNER

    [RootConfigToken("scanner")]
    [RootConfigAllowMultiple(false)]
    public class Scanner : ConfigBranch {
      private bool configured = false;
      public bool IsConfigured {
        get { return this.configured; }
        private set { this.configured = value; }
      }
      private DeviceConnectionType connType = DeviceConnectionType.Serial;
      public DeviceConnectionType ConnectionType {
        get { return this.configured ? this.connType : DeviceConnectionType.None; }
        internal set { this.connType = value; }
      }
      private int comPort;
      public int COMPort {
        get {
          if (!this.configured) { throw new ArgumentException("scanner not configured!"); }
          if (this.connType != DeviceConnectionType.Serial) { throw new ArgumentException("scanner is not serial!"); }
          return this.comPort;
        }
        internal set { this.comPort = value; }
      }
      private int tcpPort;
      public int TCPPort {
        get {
          if (!this.configured) { throw new ArgumentException("scanner not configured!"); }
          if (this.connType != DeviceConnectionType.TCP) { throw new ArgumentException("scanner is not tcp!"); }
          return this.tcpPort;
        }
        internal set { this.tcpPort = value; }
      }
      private bool p2ext = false;
      public bool P2Extension {
        get {
          if (!this.configured) { throw new ArgumentException("scanner not configured!"); }
          return this.p2ext;
        }
        internal set { this.p2ext = value; }
      }
      private int p2wait = 30;
      public int P2Wait {
        get {
          if (!this.configured) { throw new ArgumentException("scanner not configured!"); }
          return this.p2wait;
        }
        internal set { this.p2wait = value; }
      }

      protected override bool handleToken(string token, string value) {
        switch (token) {
          case "type":
            this.IsConfigured = true;
            this.ConnectionType = (DeviceConnectionType)Attributes.GenericEnumAttribute.GetEnumWith<DeviceConnectionType, ParameterConfigTokenAttribute>(value);
            return true;
          case "com":
            this.IsConfigured = true;
            this.COMPort = Convert.ToInt32(value);
            return true;

          case "port":
            this.IsConfigured = true;
            this.TCPPort = Convert.ToInt32(value);
            return true;

          case "p2ext":
            this.IsConfigured = true;
            this.P2Extension = Convert.ToBoolean(value);
            return true;
          case "p2wait":
            this.IsConfigured = true;
            this.P2Wait = Convert.ToInt32(value);
            return true;
          default:
            return false;
        }
      }
    }

    #endregion

    #region ROOT -> SCALE

    [RootConfigToken("scale")]
    [RootConfigAllowMultiple(false)]
    public class Scale : ConfigBranch {
      private bool configured = false;
      public bool IsConfigured {
        get { return this.configured; }
        private set { this.configured = value; }
      }
      private DeviceConnectionType connType = DeviceConnectionType.Serial;
      public DeviceConnectionType ConnectionType {
        get { return this.configured ? this.connType : DeviceConnectionType.None; }
        internal set { this.connType = value; }
      }
      private int comPort;
      public int COMPort {
        get {
          if (!this.configured) { throw new ArgumentException("scale not configured!"); }
          if (this.connType != DeviceConnectionType.Serial) { throw new ArgumentException("scale is not serial!"); }
          return this.comPort;
        }
        internal set { this.comPort = value; }
      }
      private ScaleModel model = ScaleModel.MettlerToledo;
      public ScaleModel Model {
        get {
          if (!this.configured) { throw new ArgumentException("scale not configured!"); }
          return this.model;
        }
        internal set { this.model = value; }
      }
      private bool softwareZero = false;
      public bool SoftwareZero {
        get {
          if (!this.configured) { throw new ArgumentException("scale not configured!"); }
          return this.softwareZero;
        }
        internal set { this.softwareZero = value; }
      }

      protected override bool handleToken(string token, string value) {
        switch (token) {
          case "type":
            this.IsConfigured = true;
            this.ConnectionType = (DeviceConnectionType)Attributes.GenericEnumAttribute.GetEnumWith<DeviceConnectionType, ParameterConfigTokenAttribute>(value);
            return true;
          case "com":
            this.IsConfigured = true;
            this.COMPort = Convert.ToInt32(value);
            return true;
          case "model":
            this.IsConfigured = true;
            this.Model = (ScaleModel)Attributes.GenericEnumAttribute.GetEnumWith<ScaleModel, ParameterConfigTokenAttribute>(value);
            return true;
          case "softzero":
            this.IsConfigured = true;
            this.SoftwareZero = Convert.ToBoolean(value);
            return true;
          default:
            return false;
        }
      }
    }

    #endregion

    #region ROOT -> DISTANCEMETER

    [RootConfigToken("distancemeter")]
    [RootConfigAllowMultiple(true)]
    public class DistanceMeterList : ConfigBranchMulti<DistanceMeter> {
      protected override bool handleToken(string token, string value) {
        if (base.wip == null) {
          base.wip = new DistanceMeter();
        }
        switch (token) {
          case "type":
            base.wip.IsConfigured = true;
            base.wip.ConnectionType = (DeviceConnectionType)Attributes.GenericEnumAttribute.GetEnumWith<DeviceConnectionType, ParameterConfigTokenAttribute>(value);
            return true;
          case "com":
            base.wip.IsConfigured = true;
            base.wip.COMPort = Convert.ToInt32(value);
            return true;
          case "model":
            base.wip.IsConfigured = true;
            base.wip.Model = (DistanceMeterModel)Attributes.GenericEnumAttribute.GetEnumWith<DistanceMeterModel, ParameterConfigTokenAttribute>(value);
            return true;
          case "history":
            base.wip.IsConfigured = true;
            base.wip.History = Convert.ToInt32(value);
            return true;
          case "stability":
            base.wip.IsConfigured = true;
            base.wip.StabilityLevel = Convert.ToInt32(value);
            return true;
          case "logging":
            base.wip.IsConfigured = true;
            base.wip.Logging = Convert.ToBoolean(value);
            return true;
          case "distancealg":
            base.wip.IsConfigured = true;
            base.wip.DistanceAlgorithm = value;
            return true;
          case "stabilityalg":
            base.wip.IsConfigured = true;
            base.wip.StabilityAlgorithm = value;
            return true;
          case "truncate":
            base.wip.IsConfigured = true;
            base.wip.TruncateCount = Convert.ToInt32(value);
            return true;
          default:
            return false;
        }
      }
    }
    public class DistanceMeter {
      private bool configured = false;
      public bool IsConfigured {
        get { return this.configured; }
        internal set { this.configured = value; }
      }
      private DeviceConnectionType connType = DeviceConnectionType.Serial;
      public DeviceConnectionType ConnectionType {
        get { return this.configured ? this.connType : DeviceConnectionType.None; }
        internal set { this.connType = value; }
      }
      private int comPort;
      public int COMPort {
        get {
          if (!this.configured) { throw new ArgumentException("distance meter not configured!"); }
          if (this.connType != DeviceConnectionType.Serial) { throw new ArgumentException("distance meter is not serial!"); }
          return this.comPort;
        }
        internal set { this.comPort = value; }
      }
      private DistanceMeterModel model = DistanceMeterModel.HRUSBMaxSonar;
      public DistanceMeterModel Model {
        get {
          if (!this.configured) { throw new ArgumentException("distance meter not configured!"); }
          return this.model;
        }
        internal set { this.model = value; }
      }
      private int history = 70;
      public int History {
        get {
          if (!this.configured) { throw new ArgumentException("distance meter not configured!"); }
          return this.history;
        }
        internal set { this.history = value; }
      }
      private int stabilityLevel = 30;
      public int StabilityLevel {
        get {
          if (!this.configured) { throw new ApplicationException("distance meter not configured!"); }
          return this.stabilityLevel;
        }
        internal set { this.stabilityLevel = value; }
      }
      private bool logging = false;
      public bool Logging {
        get {
          if (!this.configured) { throw new ApplicationException("distance meter not configured!"); }
          return this.logging;
        }
        internal set { this.logging = value; }
      }
      private string distanceAlgorithm = "mode";
      public string DistanceAlgorithm {
        get {
          if (!this.configured) { throw new ApplicationException("distance meter not configured!"); }
          return this.distanceAlgorithm;
        }
        internal set {
          if (value == "mode" || value == "median") {
            this.distanceAlgorithm = value;
          }
          else {
            throw new ArgumentException("invalid distanceAlgorithm '" + value + "'");
          }
        }
      }
      private string stabilityAlgorithm = "modeplus";
      public string StabilityAlgorithm {
        get {
          if (!this.configured) { throw new ApplicationException("distance meter not configured!"); }
          return this.stabilityAlgorithm;
        }
        internal set {
          if (value == "mode" || value == "modeplus") {
            this.stabilityAlgorithm = value;
          }
          else {
            throw new ArgumentException("invalid distanceAlgorithm '" + value + "'");
          }
        }
      }
      private int truncateCount = 0;
      public int TruncateCount {
        get {
          if (!this.configured) { throw new ApplicationException("distance meter not configured!"); }
          return this.truncateCount;
        }
        internal set {
          this.truncateCount = value;
        }
      }
    }

    #endregion

    #region ROOT -> CAMERA

    [RootConfigToken("camera")]
    [RootConfigAllowMultiple(false)]
    public class Camera : ConfigBranch {
      private bool configured = false;
      public bool IsConfigured {
        get { return this.configured; }
        private set { this.configured = value; }
      }
      private DeviceConnectionType connType = DeviceConnectionType.Serial;
      public DeviceConnectionType ConnectionType {
        get { return this.configured ? this.connType : DeviceConnectionType.None; }
        internal set { this.connType = value; }
      }

      protected override bool handleToken(string token, string value) {
        switch (token) {
          case "type":
            this.IsConfigured = true;
            this.ConnectionType = (DeviceConnectionType)Attributes.GenericEnumAttribute.GetEnumWith<DeviceConnectionType, ParameterConfigTokenAttribute>(value);
            return true;
          default:
            return false;
        }
      }
    }

    #endregion

    #region ROOT -> CUBISCAN

    [RootConfigToken("cubiscan")]
    [RootConfigAllowMultiple(false)]
    public class CubiScan : ConfigBranch {
      private bool configured = false;
      public bool IsConfigured {
        get { return this.configured; }
        private set { this.configured = value; }
      }
      private int comPort;
      public int COMPort {
        get {
          if (!this.configured) { throw new ArgumentException("cubiscan not configured!"); }
          return this.comPort;
        }
        internal set { this.comPort = value; }
      }
      private bool logging = false;
      public bool Logging {
        get {
          if (!this.configured) { throw new ApplicationException("cubiscan not configured!"); }
          return this.logging;
        }
        internal set { this.logging = value; }
      }

      protected override bool handleToken(string token, string value) {
        switch (token) {
          case "com":
            this.IsConfigured = true;
            this.COMPort = Convert.ToInt32(value);
            return true;
          case "logging":
            this.IsConfigured = true;
            this.Logging = Convert.ToBoolean(value);
            return true;
          default:
            return false;
        }
      }
    }

    #endregion

    #region ROOT -> HEATSEALER

    [RootConfigToken("heatsealer")]
    [RootConfigAllowMultiple(false)]
    public class HeatSealer : ConfigBranch {
      private bool attached = false;
      public bool Attached {
        get { return this.attached; }
        internal set { this.attached = value; }
      }

      protected override bool handleToken(string token, string value) {
        switch (token) {
          case "attached":
            this.Attached = Convert.ToBoolean(value);
            return true;
          default:
            return false;
        }
      }
    }

    #endregion

    #region ROOT -> MAS200

    [RootConfigToken("mas200")]
    [RootConfigAllowMultiple(false)]
    public class MAS200 : ConfigBranch {
      private bool forcejson = false;
      public bool ForceJSON {
        get { return this.forcejson; }
        set { this.forcejson = value; } // gets cached in TP.MAS200, so needs to be set immediately at program start!
      }

      protected override bool handleToken(string token, string value) {
        switch (token) {
          case "forcejson":
            this.ForceJSON = Convert.ToBoolean(value);
            return true;
          default:
            return false;
        }
      }
    }

    #endregion

    #region ROOT -> LAUNCHER

    [RootConfigToken("launcher")]
    [RootConfigAllowMultiple(false)]
    public class Launcher : ConfigBranch {
      private string view = "admin";
      public string View {
        get { return this.view; }
        internal set { this.view = value; }
      }

      protected override bool handleToken(string token, string value) {
        switch (token) {
          case "view":
            this.View = value; // no checking?
            return true;
          default:
            return false;
        }
      }
    }

    #endregion

    #region ROOT -> LOCATIONS

    [RootConfigToken("locations")]
    [RootConfigAllowMultiple(false)]
    public class Locations : ConfigBranch {
      private LocationDataSource force = LocationDataSource.Resource;
      public LocationDataSource Force {
        get { return this.force; }
        internal set { this.force = value; }
      }
      private string path = "c:\\";
      public string Path {
        get { return this.path; }
        internal set { this.path = value; }
      }
      private string prefix = "warehouse_";
      public string Prefix {
        get { return this.prefix; }
        internal set { this.prefix = value; }
      }
      private string extension = "osl";
      public string Extension {
        get { return this.extension; }
        internal set { this.extension = value; }
      }

      protected override bool handleToken(string token, string value) {
        switch (token) {
          case "force":
            this.Force = (LocationDataSource)Attributes.GenericEnumAttribute.GetEnumWith<LocationDataSource, ParameterConfigTokenAttribute>(value);
            return true;
          case "path":
            this.Path = value; // no checking
            return true;
          case "prefix":
            this.Prefix = value; // no checking
            return true;
          case "extension":
            this.Extension = value; // no checking
            return true;
          default:
            return false;
        }
      }

      public string Filename(int? warehouseID, int? locationTypeID, bool includeExtension) {
        return string.Format(
            "{0}{1}_{2}" + (includeExtension ? ".{3}" : ""),
            this.Prefix,
            warehouseID == null ? "A" : warehouseID.ToString(),
            locationTypeID == null ? "A" : locationTypeID.ToString(),
            this.Extension
          );
      }
    }

    #endregion

    #region ROOT -> PICKOPTS

    [RootConfigToken("pickopts")]
    [RootConfigAllowMultiple(false)]
    public class PickOptions : ConfigBranch {
      private bool fpz = false;
      public bool FPZ {
        get { return this.fpz; }
        internal set { this.fpz = value; }
      }

      protected override bool handleToken(string token, string value) {
        switch (token) {
          case "fpz":
            this.FPZ = Convert.ToBoolean(fpz);
            return true;
          default:
            return false;
        }
      }
    }

    #endregion

    #region ROOT -> IDLELOCK

    [RootConfigToken("idlelock")]
    [RootConfigAllowMultiple(false)]
    public class IdleLock : ConfigBranch {
      private bool enabled = false;
      public bool Enabled {
        get { return this.enabled; }
        internal set { this.enabled = value; }
      }
      private int timeout = 300;
      public int Timeout {
        get { return this.timeout; }
        internal set { this.timeout = value; }
      }

      protected override bool handleToken(string token, string value) {
        switch (token) {
          case "lock":
            this.Enabled = Convert.ToBoolean(value);
            return true;
          case "timeout":
            this.Timeout = Convert.ToInt32(value);
            return true;
          default:
            return false;
        }
      }
    }

    #endregion

    #region ROOT -> PRINTER

    [RootConfigToken("printer")]
    [RootConfigAllowMultiple(true)]
    public class PrinterList : ConfigBranchMulti<Printer> {
      protected override bool handleToken(string token, string value) {
        if (base.wip == null) {
          base.wip = new Printer();
        }
        switch (token) {
          case "type":
            base.wip.Type = (PrinterType)Attributes.GenericEnumAttribute.GetEnumWith<PrinterType, ParameterConfigTokenAttribute>(value);
            return true;
          case "name":
            base.wip.Name = value;
            return true;
          default:
            return false;
        }
      }

      public Printer ThermalShippingLabelPrinter() {
        return this._list.Find(delegate(Printer p) { return p.Type == PrinterType.ThermalShippingLabel; });
      }
      public Printer ThermalBarcodeLabelPrinter() {
        return this._list.Find(delegate(Printer p) { return p.Type == PrinterType.ThermalBarcode; });
      }
      public Printer RegularPrinter() {
        return this._list.Find(delegate(Printer p) { return p.Type == PrinterType.Regular; });
      }
      public Printer ImpactPrinter() {
        return this._list.Find(delegate(Printer p) { return p.Type == PrinterType.Impact; });
      }

      public string ThermalShippingLabelPrinterName() {
        Printer p = this.ThermalShippingLabelPrinter();
        return p == null ? null : p.Name;
      }
      public string ThermalBarcodeLabelPrinterName() {
        Printer p = this.ThermalBarcodeLabelPrinter();
        return p == null ? null : p.Name;
      }
      public string RegularPrinterName() {
        Printer p = this.RegularPrinter();
        return p == null ? null : p.Name;
      }
      public string ImpactPrinterName() {
        Printer p = this.ImpactPrinter();
        return p == null ? null : p.Name;
      }
    }
    public class Printer {
      private PrinterType type;
      public PrinterType Type {
        get { return this.type; }
        internal set { this.type = value; }
      }
      private string name;
      public string Name {
        get { return this.name; }
        internal set { this.name = value.Replace('+', ' '); }
      }
    }

    #endregion

    #region ROOT -> TEXT TO SPEECH

    [RootConfigToken("speech")]
    [RootConfigAllowMultiple(false)]
    public class TextToSpeech : ConfigBranch {
      private SpeechOption force = SpeechOption.Any;
      public SpeechOption Force {
        get { return this.force; }
        set { this.force = value; }
      }

      protected override bool handleToken(string token, string value) {
        switch (token) {
          case "force":
            this.Force = (SpeechOption)Attributes.GenericEnumAttribute.GetEnumWith<SpeechOption, ParameterConfigTokenAttribute>(value);
            return true;
          default:
            return false;
        }
      }
    }

    #endregion

    #region ROOT -> POD

    [RootConfigToken("pod")]
    [RootConfigAllowMultiple(false)]
    public class Pod : ConfigBranch {
      private bool configured = false;
      public bool IsConfigured {
        get { return this.configured; }
        private set { this.configured = value; }
      }
      private int address;
      public int Address {
        get {
          if (!this.configured) { throw new ArgumentException("pod not configured!"); }
          return this.address;
        }
        internal set { this.address = value; this.configured = true; }
      }
      private int comPort;
      public int COMPort {
        get {
          if (!this.configured) { throw new ArgumentException("pod not configured!"); }
          return this.comPort;
        }
        internal set { this.comPort = value; this.configured = true; }
      }
      private CarouselOrientation orientation;
      public CarouselOrientation Orientation {
        get {
          if (!this.configured) { throw new ArgumentException("pod not configured!"); }
          return this.orientation;
        }
        internal set { this.orientation = value; this.configured = true; }
      }

      protected override bool handleToken(string token, string value) {
        switch (token) {
          case "address":
            this.Address = Convert.ToInt32(value);
            return true;
          case "com":
            this.COMPort = Convert.ToInt32(value);
            return true;
          case "type":
            this.Orientation = (CarouselOrientation)Attributes.GenericEnumAttribute.GetEnumWith<CarouselOrientation, ParameterConfigTokenAttribute>(value);
            return true;
          default:
            return false;
        }
      }
    }

    #endregion

    #region ROOT -> CAROUSEL

    [RootConfigToken("carousel")]
    [RootConfigAllowMultiple(true)]
    public class CarouselList : ConfigBranchMulti<Carousel> {
      protected override bool handleToken(string token, string value) {
        if (this.wip == null) {
          this.wip = new Carousel();
        }
        switch (token) {
          case "address":
            this.wip.Address = Convert.ToInt32(value);
            return true;
          case "bins":
            this.wip.Bins = Convert.ToInt32(value);
            return true;
          default:
            return false;
        }
      }
    }
    public class Carousel {
      private int address;
      public int Address {
        get { return this.address; }
        internal set { this.address = value; }
      }
      private int bins = 36;
      public int Bins {
        get { return this.bins; }
        internal set { this.bins = value; }
      }
    }

    #endregion

    #region ROOT -> LIGHTREE
    
    [RootConfigToken("lightree")]
    [RootConfigAllowMultiple(true)]
    public class LighTreeList : ConfigBranchMulti<LighTree> {
      protected override bool handleToken(string token, string value) {
        if (this.wip == null) {
          this.wip = new LighTree();
        }
        switch (token) {
          case "address":
            this.wip.Address = Convert.ToInt32(value);
            return true;
          case "lights":
            this.wip.Lights = Convert.ToInt32(value);
            return true;
          case "chars":
            this.wip.Chars = Convert.ToInt32(value);
            return true;
          default:
            return false;
        }
      }
    }
    public class LighTree {
      private int address;
      public int Address {
        get { return this.address; }
        internal set { this.address = value; }
      }
      private int lights;
      public int Lights {
        get { return this.lights = 12; }
        internal set { this.lights = value; }
      }
      private int chars;
      public int Chars {
        get { return this.chars = 6; }
        internal set { this.chars = value; }
      }
    }

    #endregion

    #region ROOT -> SORTBAR

    [RootConfigToken("sortbar")]
    [RootConfigAllowMultiple(true)]
    public class SortBarList : ConfigBranchMulti<SortBar> {
      protected override bool handleToken(string token, string value) {
        if (this.wip == null) {
          this.wip = new SortBar();
        }
        switch (token) {
          case "address":
            this.wip.Address = Convert.ToInt32(value);
            return true;
          case "lights":
            this.wip.Lights = Convert.ToInt32(value);
            return true;
          case "chars":
            this.wip.Chars = Convert.ToInt32(value);
            return true;
          default:
            return false;
        }
      }
    }
    public class SortBar {
      private int address;
      public int Address {
        get { return this.address; }
        internal set { this.address = value; }
      }
      private int lights = 6;
      public int Lights {
        get { return this.lights; }
        internal set { this.lights = value; }
      }
      private int chars = 6;
      public int Chars {
        get { return this.chars; }
        internal set { this.chars = value; }
      }
    }

    #endregion

    #region ROOT -> CHARACTERSTRIP

    [RootConfigToken("characterstrip")]
    [RootConfigAllowMultiple(true)]
    public class CharacterStripList : ConfigBranchMulti<CharacterStrip> {
      protected override bool handleToken(string token, string value) {
        if (this.wip == null) {
          this.wip = new CharacterStrip();
        }
        switch (token) {
          case "address":
            this.wip.Address = Convert.ToInt32(value);
            return true;
          case "lights":
            this.wip.Lights = Convert.ToInt32(value);
            return true;
          case "chars":
            this.wip.Chars = Convert.ToInt32(value);
            return true;
          default:
            return false;
        }
      }
    }
    public class CharacterStrip {
      private int address;
      public int Address {
        get { return this.address; }
        internal set { this.address = value; }
      }
      private int lights = 28;
      public int Lights {
        get { return this.lights; }
        internal set { this.lights = value; }
      }
      private int chars = 1;
      public int Chars {
        get { return this.chars; }
        internal set { this.chars = value; }
      }
    }

    #endregion

    #region ROOT -> SURFACEDISPLAY

    [RootConfigToken("surfacedisplay")]
    [RootConfigAllowMultiple(true)]
    public class SurfaceDisplayList : ConfigBranchMulti<SurfaceDisplay> {
      protected override bool handleToken(string token, string value) {
        if (this.wip == null) {
          this.wip = new SurfaceDisplay();
        }
        switch (token) {
          case "address":
            this.wip.Address = Convert.ToInt32(value);
            return true;
          case "lights":
            this.wip.Lights = Convert.ToInt32(value);
            return true;
          case "chars":
            this.wip.Chars = Convert.ToInt32(value);
            return true;
          default:
            return false;
        }
      }
    }
    public class SurfaceDisplay {
      private int address;
      public int Address {
        get { return this.address; }
        internal set { this.address = value; }
      }
      private int lights = 1;
      public int Lights {
        get { return this.lights; }
        internal set { this.lights = value; }
      }
      private int chars = 10;
      public int Chars {
        get { return this.chars; }
        internal set { this.chars = value; }
      }
    }

    #endregion

    #region ROOT -> LINKAGE

    [RootConfigToken("linkage")]
    [RootConfigAllowMultiple(true)]
    public class PodLinkageList : ConfigBranchMulti<PodLinkage> {
      protected override bool handleToken(string token, string value) {
        if (this.wip == null) {
          this.wip = new PodLinkage();
        }
        switch (token) {
          case "carousel":
            this.wip.Carousel = Convert.ToInt32(value);
            return true;
          case "lighttype":
            this.wip.LightType = (LightType)Attributes.GenericEnumAttribute.GetEnumWith<LightType, ParameterConfigTokenAttribute>(value);
            return true;
          case "lightaddr":
            this.wip.LightAddr = Convert.ToInt32(value);
            return true;
          case "direction":
            this.wip.Direction = (Direction)Attributes.GenericEnumAttribute.GetEnumWith<Direction, ParameterConfigTokenAttribute>(value);
            return true;
          default:
            return false;
        }
      }
    }
    public class PodLinkage {
      private int carousel;
      public int Carousel {
        get { return this.carousel; }
        internal set { this.carousel = value; }
      }
      private LightType lighttype;
      public LightType LightType {
        get { return this.lighttype; }
        internal set { this.lighttype = value; }
      }
      private int lightaddr;
      public int LightAddr {
        get { return this.lightaddr; }
        internal set { this.lightaddr = value; }
      }
      private Direction direction; // direction, from pov of the tree
      public Direction Direction {
        get { return this.direction; }
        internal set { this.direction = value; }
      }
    }

    #endregion





    public enum DatabaseConnectionMethod {
      [ParameterConfigToken("sspi")]
      SSPI,
      [ParameterConfigToken("userpass")]
      UserPass,
    };

    public enum DeviceConnectionType {
      [ParameterConfigToken("none")]
      None,
      [ParameterConfigToken("serial")]
      Serial,
      [ParameterConfigToken("usb")]
      USB,
      [ParameterConfigToken("tcp")]
      TCP,
    };

    public enum ScaleModel {
      [ParameterConfigToken("none")]
      None,
      [ParameterConfigToken("fairbanks")]
      Fairbanks,
      [ParameterConfigToken("mettlertoledo")]
      MettlerToledo,
    };

    public enum DistanceMeterModel {
      [ParameterConfigToken("none")]
      None,
      [ParameterConfigToken("hrusb-maxsonar")]
      HRUSBMaxSonar,
    };

    public enum PrinterType {
      [ParameterConfigToken("regular")]
      Regular,
      [ParameterConfigToken("thermal-barcode")]
      ThermalBarcode,
      [ParameterConfigToken("thermal-shiplabel")]
      ThermalShippingLabel,
      [ParameterConfigToken("impact")]
      Impact,
    };

    public enum LocationDataSource {
      [ParameterConfigToken("database")]
      Database,
      [ParameterConfigToken("file")]
      File,
      [ParameterConfigToken("resource")]
      Resource,
    };

    public enum LightType {
      [ParameterConfigToken("lightree")]
      LighTree,
      [ParameterConfigToken("characterstrip")]
      CharacterStrip,
      [ParameterConfigToken("surfacedisplay")]
      SurfaceDisplay,
    };

    public enum CarouselOrientation {
      [ParameterConfigToken("horizontal")]
      Horizontal,
      [ParameterConfigToken("vertical")]
      Vertical,
    };

    public enum Direction {
      [ParameterConfigToken("left")]
      Left,
      [ParameterConfigToken("right")]
      Right,
      [ParameterConfigToken("unknown")]
      Unknown,
    };

    public enum SpeechOption {
      [ParameterConfigToken("any")]
      Any,
      [ParameterConfigToken("api")]
      API,
      [ParameterConfigToken("builtin")]
      BuiltIn,
      [ParameterConfigToken("off")]
      Off,
    };

  }
}
