        ��  ��                        �� ��     0         <?xml version="1.0" encoding="UTF-8"?>
<assembly xmlns="urn:schemas-microsoft-com:asm.v1" manifestVersion="1.0">
  <assemblyIdentity
    version="1.0.0.16"
    processorArchitecture="X86"
    name="FLExible Patcher"
    type="win32"
    />
  <description></description>
    <dependency>
        <dependentAssembly>
            <assemblyIdentity
                type="win32"
                name="Microsoft.Windows.Common-Controls"
                version="6.0.0.0"
                processorArchitecture="X86"
                publicKeyToken="6595b64144ccf1df"
                language="*"
             />
        </dependentAssembly>
    </dependency>
<!-- Identify the application security requirements: Vista and above -->
  <trustInfo xmlns="urn:schemas-microsoft-com:asm.v2">
      <security>
        <requestedPrivileges>
          <requestedExecutionLevel
            level="asInvoker"
            uiAccess="false"
            />
        </requestedPrivileges>
      </security>
  </trustInfo>
</assembly>  �   ,   C U S T O M   ��h       0        patchsetversion=1.0
  header=0
      enabled=1
      name=Empty custom patch
      description=
      category=
      dependencies=
      filename=tomb4.exe
  footer=0
 z  ,   C U S T O M   ��g       0        presetversion=1.0
 silentpatch=0
 developermode=0
 developerview=0
 defaultpatchset=patches.flp
 maxparameters=30
 defaultexename=tomb4.exe
 lastpreset=
 windowtitle=FLEP
 abouttext1=FLEP - version 1.1.//Based on TREP source code.
 abouttext2=Source code by Ensi.//Modified and refactored by Lwmte.
 abouttext3=Listen to the flowers grow.//God is an astronaut.
  