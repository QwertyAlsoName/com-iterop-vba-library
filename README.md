# com-iterop-vba-library

一个 .NET COM 互操作库，用于将计算函数暴露给 VBA/Excel 宏。使用 C# 编写，目标框架为 .NET Framework 3.5，注册为 COM 服务器后，Excel 工作簿可通过 `CreateObject("vbdl")` 或早期绑定（通过生成的类型库）调用其功能。

## 命名空间

`Iterop.VbaLibrary`

## 类

### `VbaHelper`（ProgId：`vbdl`）

用于连通性测试的工具类。

| 方法 | 说明 |
|------|------|
| `Hello()` | 返回字符串 `"Hello"`，用于验证 COM 注册是否成功。 |

### `DotNetCalc`

暴露油藏工程计算方法。

| 方法 | 说明 |
|------|------|
| `Sgr(porosity, bubblePressureKPa, temperatureCelsius, gorAtBubblePoint, waterCompressibility, oilSpecificGravity, minimumPressureKPa)` | 计算地层条件下的溶解气油比（SGR）。接受公制/SI 单位输入，内部转换为油田单位（psi、°F、scf/bbl、API）。 |

**`Sgr` 参数说明：**

| 参数 | 单位 | 说明 |
|------|------|------|
| `porosity` | 无量纲（0–1） | 地层孔隙度 |
| `bubblePressureKPa` | kPa | 泡点压力 |
| `temperatureCelsius` | °C | 地层温度 |
| `gorAtBubblePoint` | m³/m³ | 泡点处溶解气油比 |
| `waterCompressibility` | 1/kPa × 10⁻⁶ | 地层水压缩系数 |
| `oilSpecificGravity` | 无量纲 | 原油相对密度（必须大于零） |
| `minimumPressureKPa` | kPa | 最低地层压力（必须大于零） |

## 环境要求

- Windows 操作系统
- .NET Framework 3.5 或更高版本
- 管理员权限（用于 COM 注册）

## 构建

在 Visual Studio 或 JetBrains Rider 中打开 `com-iterop-vba-library.sln`，选择 **Release** 配置进行构建。程序集使用 `key_my_com.pfx` 进行强名称签名。

## 注册

构建完成后以管理员身份运行以下命令（或使用 `_Test\` 目录中的脚本）：

```bat
C:\Windows\Microsoft.NET\Framework\v4.0.30319\RegAsm.exe "path\to\vbdl.dll" /tlb /codebase
```

注销命令：

```bat
C:\Windows\Microsoft.NET\Framework\v4.0.30319\RegAsm.exe /u "path\to\vbdl.dll"
```

`_Test\register.bat` 和 `_Test\unregister.bat` 脚本已针对测试构建自动化上述操作。

## 在 VBA 中使用

```vba
' 连通性测试
Dim helper As Object
Set helper = CreateObject("vbdl")
MsgBox helper.Hello()

' 溶解气油比计算
Dim calc As Object
Set calc = CreateObject("vbdl.DotNetCalc")
Dim result As Double
result = calc.Sgr(porosity, bubblePressureKPa, temperatureCelsius, _
                  gorAtBubblePoint, waterCompressibility, oilSpecificGravity, _
                  minimumPressureKPa)
```

`_Test\` 目录中提供了测试工作簿（`Testing2.xlsm`、`Testing2.xlsx`）。

## 项目结构

```
com-iterop-vba-library.sln
com-iterop-vba-library/
  Main.cs                  — COM 类（VbaHelper、DotNetCalc）
  Properties/
    AssemblyInfo.cs
_Test/
  vbdl.dll                 — 预编译测试二进制文件
  Testing2.xlsm            — Excel 测试工作簿（含宏）
  Testing2.xlsx            — Excel 测试工作簿
  register.bat             — 注册 DLL
  unregister.bat           — 注销 DLL
```

## 版权

© 2019 Iterop．商标所有人：王心月
