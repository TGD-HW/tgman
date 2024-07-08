**TGMcontroller** je hardware pro kompletní řídicí systém **TGMotion**. 
Podrobný popis systému TGMotion je uveden v jeho vlastní příručce.
Mezi TGMotion na PC a v TGMcontrolleru nejsou žádné funkční rozdíly.
Virtuální PLC lze v případě potřeby kompletně vyvinout a otestovat na PC a poté pouze zkompilovat pro TGMcontroller a spustit na něm bez dalších změn.

##Control observer
Control Observer je servisní a monitorovací nástroj běžící na počítači s Windows a hraje velkou roli při údržbě, aktualizaci firmwaru, vývoji PLC a uvádění řídicí jednotky TGM do provozu.
Nejprve musí být připojen k zařízení. K tomu slouží okno *Connection info* :

![Connection info dialog](../img/connectionInfo.png){: style="width:60%;" }

V menu `Connection Device` zvolte požadovaný typ zařízení.
Použijte připojení TCP nebo UDP, pokud je IP adresa již známá (zobrazuje se při spuštění na LED displeji) a je platná v rámci ethernetového segmentu.
Použijte typ FSP, pokud je ethernetový kabel připojen z PC přímo (peer to peer) ke konektoru FSP (X12).
Počkejte, až bude stav připojení ONLINE. Pak jsou k dispozici všechny funkce a **Control Observer** pracuje přímo se zařízením.

##Konfigurační soubor `TGM.INI`
Soubor `TGM.INI` je uložen na SD kartě.
Lze jej editovat přímo na kartě pomocí počítače nebo jej lze přečíst ze zařízení pomocí Control Observeru, upravit a poté zapsat zpět.
Použijte okno *Systémové soubory*, upravte odpovídajícím způsobem názvy v editačních polích a použijte tlačítka `[Read]` a poté `[Write]` (řádky souboru INI).
Názvy souborů v počítači mohou být libovolné, TGMcontroller interně používá správné názvy souborů SD karty.

![System files dialog](../img/systemFiles.png){: style="width:60%;" }

Po úpravě souboru `TGM.INI` je nutné restartovat TGMcontroller (tlačítkem `[Restart Device]` nebo vypnutím/zapnutím zařízení).
Vnitřní zpětné vazby a vstupy/výstupy musí být v případě potřeby namapovány na TGMotion v souboru `TGM.INI`. 
Hodnoty jsou viditelné ve struktuře `SERVO`. Použijte typ serva 92:

``` ini
[Servo_Configuration]

Servo[00].Type=92
Servo[00].Node=1
Servo[00].Axis=1
Servo[00].FeedbackType=1

Servo[01].Type=92
Servo[01].Node=1
Servo[01].Axis=2
Servo[01].FeedbackType=1
```

Hodnota `Node` je irelevantní a ignoruje se, ale měla by být v souboru INI jedinečná a v rozsahu 1 - 64.
`Axis=1` podporuje `FeedbackType 1` (DSL) a `2` (EnDAT), zatímco `Axis=2` podporuje `FeedbackType 1` (DSL) a `3` (SSI).
Jiné hodnoty jsou neplatné.
Pokud není typ zpětné vazby uveden, použije se hodnota `1` (DSL).
Indexy v hranatých závorkách (00 nebo 01 výše) mohou být samozřejmě různé, v rozsahu 0 - 63.   

Skutečné hodnoty polohy z feedbacku se ukládají do proměnné `SERVO[xx]`.
Proměnná `Position` (sdílená paměť TGMotion SERVO, přístupná z PLC nebo PC), kde `"xx"` je index z výše uvedeného souboru `TGM.INI`.
Hodnota z inkrementálního snímače (konektor X5) je vždy namapována na osu 1, proměnná `SERVO[xx].ExtPosition`.
Stejně tak se digitální vstupy a výstupy nacházejí v proměnných `SERVO[xx].DigitalIn` resp. `SERVO[xx].DigitalOut`.
V následující tabulce je uvedeno mapování použitých vstupů.
Předpokládá výše uvedené záznamy v souboru `TGM.INI` (`SERVO[00].Axis=1` a `SERVO[01].Axis=2`).

| Vstupní pin| Osa | Výskyt u bitu | Příklad kódu               |
|--------------|-----|------------------|----------------------------|
| DI1           | 1   | DigitalIn.0      | SERVO[00].DigitalIn & 0x001 |
| DI2           | 2   | DigitalIn.0      | SERVO[01].DigitalIn & 0x001 |
| DI3           | 1   | DigitalIn.1      | SERVO[00].DigitalIn & 0x002 |
| DI4           | 2   | DigitalIn.1      | SERVO[01].DigitalIn & 0x002 |
| DI5           | 1   | DigitalIn.2      | SERVO[00].DigitalIn & 0x004 |
| DI6           | 2   | DigitalIn.2      | SERVO[01].DigitalIn & 0x004 |
| DI7           | 1   | DigitalIn.3      | SERVO[00].DigitalIn & 0x008 |
| DI8           | 2   | DigitalIn.3      | SERVO[01].DigitalIn & 0x008 |
| EI1           | 1   | DigitalIn.8      | SERVO[00].DigitalIn & 0x100 |
| EI2           | 1   | DigitalIn.9      | SERVO[00].DigitalIn & 0x200 |

Mapování výstupů je podobné:

| Výstupní pin | Osa | Použitý bit  | Kód pro nastavení do '1'                    | Kód pro nastavení do '0'                       |
|-----------------|----|---------------|---------------------------------|-------------------------------------|
| DO1                | 1   | DigitalOut.0  | SERVO[00].DigitalOut \|= 0x01  | SERVO[00].DigitalOut &= ~0x01       |
| DO2                | 2   | DigitalOut.0  | SERVO[01].DigitalOut \|= 0x01  | SERVO[01].DigitalOut &= ~0x01       |
| DO3                | 1   | DigitalOut.1  | SERVO[00].DigitalOut \|= 0x02  | SERVO[00].DigitalOut &= ~0x03       |
| DO4                | 2   | DigitalOut.1  | SERVO[01].DigitalOut \|= 0x02  | SERVO[01].DigitalOut &= ~0x03       |
| DO5                | 1   | DigitalOut.2  | SERVO[00].DigitalOut \|= 0x04  | SERVO[00].DigitalOut &= ~0x04       |
| DO6                | 2   | DigitalOut.2  | SERVO[01].DigitalOut \|= 0x04  | SERVO[01].DigitalOut &= ~0x04       |

Hodnoty analogových vstupů AN1, AN2 se objeví v proměnných `SERVO[00].AnalogIn` a `SERVO[01].AnalogIn`.


##Komunikace s počítačem
TGMcontroller používá dva gigabitové ethernetové porty pro komunikaci s okolním světem.
Tzv. servisní port (X11) implementuje standardní protokoly: TCP, UDP, Profinet I/O a Modbus TCP.
Všechny tyto protokoly lze v případě potřeby používat současně.
Důležitou roli pro tyto protokoly hraje IP adresa (viz níže).
Tento servisní port lze použít ve větších sítích, kde jsou nutné ethernetové switche, směrovače nebo bridge.
Současně může být připojeno až 16 externích klientů.   

Protože TGMcontroller lze použít také jako náhradu za rozšíření PC v reálném čase, je implementován Fast Service Port (X12).
Používá vlastní protokol a měl by být používán pouze pro peer to peer ethernetové připojení, tj. přímý kabel mezi PC a TGMcontrollerem.
FSP se může vyznačovat velmi rychlou dobou odezvy a umožňuje velmi vysokou šířku pásma.
Nepoužívá žádnou IP adresu, a proto se velmi snadno nastavuje.
Stačí připojit ethernetový kabel k portu X12 FSP a k libovolnému volnému adaptéru síťové karty Ethernet v počítači (nejlépe k vestavěnému adaptéru na základní desce nebo k adaptéru PCIe).
Podpůrná knihovna DLL komunikace Control Observer (`TGM_DEV_FSP_5.dll` nebo `TGM_DEV_FSP_5_x64.dll`) prohledá všechny dostupné adaptéry PC a při spuštění vybere ten připojený.
Nevýhodou je, že v počítači musí být nainstalován ovladač [Winpcap](https://www.winpcap.org/) nebo [Npcap](https://npcap.com/).

##IP adresa
Výchozí IP adresa je `192.168.1.188`. Lze ji změnit v souboru `TGM.INI`:

``` ini
[DHCP]
Enable=0
IPAddress=192.168.1.188
Mask=255.255.255.0
Gateway=192.168.1.1
DNS=8.8.8.8
```
Po změně IP adresy nezapomeňte ve stejném ethernetovém segmentu nastavit také správnou adresu brány.   

Pro automatické přidělování IP adres je také možné použít DHCP (v sítích se směrovačem).
Nastavením `Enable=1` se tato funkce aktivuje.
Když proces DHCP proběhne úspěšně, na LED displeji se zobrazí přidělená IP adresa.
Upozorňujeme, že to může nějakou dobu trvat.
Pokud se proces DHCP nezdaří (po uplynutí časového limitu několika desítek sekund), použije se místo toho statická IP adresa ze souboru `TGM.INI`.

##Adresa MAC
TGMcontroller přečte svou interní jedinečnou hodnotu DNA a vytvoří z ní adresu MAC.
Automatická adresa MAC začíná vždy `00:0A`. Tuto hodnotu lze přepsat záznamem v souboru TGM.INI v sekci [DHCP]:

``` ini
[DHCP]
...
MacAddress=00.xx.xx.xx.xx.xx.xx
```

První číslo musí být vždy `00`, jinak je položka ignorována.
Ostatní hodnoty `"xx"` mohou být libovolné.
Všimněte si, že se používají hexadecimální čísla.

##Vývoj a nahrávání PLC
Zdrojový kód PLC TGMotion je vždy kompatibilní se všemi deriváty TGMotion (Windows PC s rozšířením pro reálný čas, TGMmini, TGZ+Motion, TGMcontroller a případnými budoucími systémy).
To znamená, že PLC lze naprogramovat a otestovat např. na PC a poté pouze překompilovat pro systém TGMcontroller. PLC musí být napsáno v jazyce C/C++.   

Program PLC lze zkompilovat na libovolném počítači se systémem Windows nebo Linux.
Všechny potřebné aplikace jsou open source.
Je nutný meziplatformní kompilátor.
Musí to být kompilátor `GCC` s názvem `gcc-arm-none-eabi`.
Nejnovější verzi lze stáhnout z [webových stránek ARM](https://developer.arm.com/).
Vyberte nejnovější verzi ZIP a rozbalte ji do libovolné složky.
Překladač musí podporovat alespoň standard C++17.
Projekt pro PLC používá standardní makefile GNU.
K provedení příkazů `makefile` je nutné použít program `mingw32-make`.
Program `mingw32-make` lze stáhnout např. ze [sourceforge.net](https://sourceforge.net/projects/mingw/) a musí být dostupný z příkazového řádku (např. přidáním jeho cesty do systémové proměnné `PATH` nebo použitím jeho názvu spolu s cestou v příkazovém řádku).
Kompilátor GCC není nutné mít v `PATH`.
Makefile má téměř pevnou strukturu, pouze seznam .c/.h souborů a cestu ke kompilátoru je nutné upravovit.
Pokud je makefile správný, lze PLC vytvořit příkazem

``` powershell
mingw32-make all
```

v adresáři, kde se nachází soubor makefile a zdrojové soubory.
Volitelně použijte úplnou cestu k souboru mingw32-make.exe, např.

``` powershell
C:\PLC\projects\gnu\mingw32-make all
```

Chcete-li znovu sestavit PLC, nejprve vyčistěte mezisoubory pomocí následujících kroků. 

``` powershell
mingw32-make clean
```

a pak jej sestavte pomocí

``` powershell
mingw32-make all
```

Chcete-li jej zkompilovat s využitím více jader procesoru, použijte volbu `-j<číslo>`, např.

``` powershell
mingw32-make -j4 all
```
(používá čtyři vlákna pro kompilaci projektu, což urychluje proces).   

Příklad souboru makefile:

``` mf
#path to TGM include files
TGM_PATH=..\include
#path to compiler
CC_PATH=..\..\gnu\gcc-arm-none-eabi\bin
#source files, must be in the same directory as this makefile
SOURCES=main.cpp \
DI_Capture.cpp \
Program_01.cpp \
Program_02.cpp \
Program_03.cpp \
Program_04.cpp \
Program_Ini.cpp \
Servo_Func.cpp \
$(TGM_PATH)\fmt\format.cpp
#used header file. these files are checked for dependencies
HEADERS=Definition.h \
DI_Capture.h \
PLC_Func.h \
Servo_Func.h \
stdafx.h \
User_Definitions.h \
User_Variables.h
USER_PATH=.\
#set the output filename
OUT_FILE_NAME=TGM_Controller_CAN_Tester
#===============================
#there is no need to change the rest of makefile
#compiler file names
CC=$(CC_PATH)\arm-none-eabi-g++.exe
OBJCOPY=$(CC_PATH)\arm-none-eabi-objcopy.exe
OBJDUMP=$(CC_PATH)\arm-none-eabi-objdump.exe
#stamper is special utility to mark the resulting PLC as for TGMcontroller
STAMPER=tgm_bin_stamper.exe
OPTIMIZE=-O2
BSP_PATH=..\BSP
EXECUTABLE=$(OUT_FILE_NAME).elf
BIN_FILE=$(OUT_FILE_NAME).tgm.bin
STAMPED_FILE=$(OUT_FILE_NAME).tgm_controller
DUMP_FILE=$(OUT_FILE_NAME).elf.dbg
MAX_SERVO=64
MAX_DIO=16
MAX_INTERPOLATOR=3
SIZE_OSCILLOSCOPE=33554432
TGM_SETTINGS=-DMAX_SERVO_PROJECT_SETTINGS=$(MAX_SERVO) \
-DMAX_DIO_PROJECT_SETTINGS=$(MAX_DIO) \
-DMAX_INTERPOLATOR_PROJECT_SETTINGS=$(MAX_INTERPOLATOR) \
-DSIZE_OSCILLOSCOPE_MEMORY_PROJECT_SETTINGS=$(SIZE_OSCILLOSCOPE)
DEFINES=$(TGM_SETTINGS) \
-DFREERTOS \
-DZYNQ \
-DNDEBUG \
-DGLOBAL_TIMER_PRESCALER=1 \
-DNO_FILE_FUNCTIONS
CFLAGS=$(OPTIMIZE) -std=c++17 -I. -I$(BSP_PATH)\include -I$(TGM_PATH) \
-I$(USER_PATH) $(DEFINES) -g -Wall -Wextra -mcpu=cortex-a9 \
-mfpu=vfpv3 -mfloat-abi=hard -ffunction-sections -fdata-sections \
-Wno-unknown-pragmas
LDFLAGS=-g -mlittle-endian -Wl,-T -Wl,lscript.ld \
-Wl,--start-group,-lxil,-lgcc,-lc,-lstdc++,--end-group \
-L$(BSP_PATH)\lib -mcpu=cortex-a9 -mfpu=vfpv3 -mfloat-abi=hard \
-Wl,-build-id=none -specs=Xilinx.spec -Wl,--gc-sections
OBJECTS=$(SOURCES:.cpp=.o)
$(EXECUTABLE): $(OBJECTS) $(SOURCES) $(HEADERS) makefile
$(CC) $(OBJECTS) $(LIB_OBJECTS_C) $(LIB_OBJECTS_CPP) $(LIB_OBJECTS_ASM) \
-o $@ $(LDFLAGS)
cmd /C $(OBJDUMP) -C -x -S $(EXECUTABLE) >$(DUMP_FILE)
$(OBJCOPY) -O binary $(EXECUTABLE) $(BIN_FILE)
$(STAMPER) $(BIN_FILE) $(STAMPED_FILE) -oPLC -b0x07800000 -s0x00800000
%.o: %.cpp $(HEADERS) makefile | $(OBJDIR)
$(CC) -c $< $(CFLAGS) -o $@
all: $(SOURCES) $(HEADERS) $(EXECUTABLE) makefile
clean:
del /Q $(OBJECTS) $(EXECUTABLE) $(BIN_FILE) $(DUMP_FILE)
```

Po vytvoření spustitelného souboru `.ELF` vytvoří soubor makefile binární soubor a následně jej orazítkuje samostatným nástrojem `tgm_bin_stamper.exe`.   

Požádejte zástupce společnosti TG Drives o příklad programu PLC.
Na SD kartě je také příklad programu spolu s dalšími pomocnými soubory.
Projekt obsahuje všechny potřebné soubory, nástroje kompilátoru, soubor linkerového skriptu, program stamper, BSP (board support package) a veřejné soubory TGMotion include.
Soubor `main.cpp` obsahuje nezbytnou přístupovou funkci `GetProcAddress()`, která slouží k získání adres umístění hlavních funkcí PLC: `Program_Ini(), Program_01(), Program_02(), Program_03(), Program_04()`.
Funkce `GetProcAddress()` musí být umístěna na startovací adrese `0x07800000` a nesmí být optimalizována.
Dále inicializuje standardní knihovnu jazyka C a všechny závislosti, jako jsou ukazatele na virtuální funkce a inicializace lokálních a globálních proměnných.   

Viz. také samostatná příručka o programování PLC pomocí TGMotion.


##Spuštění PLC a autostart
Kromě okna *Systémové soubory* programu **Control Observer** je k dispozici okno určené pro programování PLC - *PLC Control*.
Stačí nastavit správný název souboru PLC a nahrát jej do **TGMcontroller**.
PLC se spustí kliknutím na tlačítko `[Spustit]`. 

![PLC control dialog](../img/PLCcontrol.png){: style="width:60%;" }

Proces `Run PLC` provede následující sekvenci:

1. Zastaví všechny běžící PLC a počká na dokončení všech funkcí `Program_XX()`.
2. Vymaže paměť PLC DATA.
3. Vymaže (nastaví na nulu) všechny digitální výstupy připojených I/O zařízení i servopohonů. Nastaví režim všech servopohonů na nulu.
4. Nastaví také interní digitální výstupy na nulu (pokud jsou namapovány na TGMotion).
5. Hlavní smyčka serva je zastavena a nejsou přenášeny žádné další zprávy EtherCAT.
6. PLC se načte do paměti z karty SD.
7. Hlavní smyčka serva se znovu spustí, komunikace EtherCAT se aktivuje.
8. Zavolá se inicializační funkce PLC, `Program_Ini()`, a pokud je úspěšná (vrátí 1), zavede se periodické volání funkce `Program_XX()`, tj. PLC se spustí.

   
	
PLC lze spustit po spuštění TGMcontrolleru.
Nastavte položku v `TGM.INI`

``` ini
[PLC_Configuration]
…
Autostart_PLC=1
```

Mezi spuštěním PLC a komunikací EtherCAT není žádná synchronizace.
Obvykle se PLC spouští rychleji než komunikace EtherCAT.
Programátor s tím musí počítat a čekat například na hodnotu stavu EtherCAT (`operational = 8`) připojených zařízení nebo na požadovaný počet nalezených zařízení EtherCAT na sběrnici (proměnná `SYSTEM.MAIN.Found_Total_Number_Of_ECAT_Devices`).

##Priorita programů PLC a doba cyklu
Programy v PLC běží s různou prioritou.
Je na programátorovi, aby zvolil správnou funkci pro úlohy PLC.

1.  `Program_04()` má nejvyšší prioritu a probíhá v kontextu funkce servocyklu.
	Je volán periodicky v intervalu daném hodnotou `Cycle_Time` v souboru `TGM.INI` (v mikrosekundách).
	Přípustné hodnoty jsou 100, 200, 250, 500, 1000, 2000 a pak 3000, ..., 10000 (po 1000 krocích).
	Je velmi důležité, aby `Program_04()` netrval déle než přibližně polovinu doby cyklu, jinak by TGMotion neměl šanci provést vlastní výpočty.
	Příliš zaneprázdněná funkce `Program_04()` není kontrolována, ale problém se velmi rychle objeví na straně EtherCAT systému: 
	požadované polohy servopohonů nebudou aktualizovány a v samotném servozesilovači vznikne následující chyba.
	Pokud je bezpodmínečně nutné provést náročný výpočet ve funkci `Program_04()`, je jediným řešením prodloužení doby cyklu.
		
2. `Program_03()` má prioritu jen o jeden stupeň nižší, než `Program_04()` a podobně `Program_02()` o dva stupně nižší, než `Program_04()`.
	Programy jsou volány opakovaně s dobou cyklu uvedenou v souboru `TGM.INI`. Například:
		
	``` ini
	[PLC_Configuration]
	...
	Cycle_Time_Program_02=400
	Cycle_Time_Program_03=800
	```
		
	Hodnoty se udávají v mikrosekundách a musí být násobkem 200.
	Minimální povolená hodnota je 200&nbsp;µs.
	Chcete-li danou funkci zcela zakázat, použijte hodnotu nula.
	TGMcontroller kontroluje uplynulý čas v každé funkci, a pokud do dalšího volání funkce zbývá méně než 100&nbsp;mikrosekund, ke skutečnému volání dojde až při dalším tiku 200 mikrosekund.
	TGMcontroller tak může naplánovat jiné úlohy s nižší prioritou (komunikace atd.) a `Program_02()` nebo `Program_03()` nemůže přetížit systém.
		
3. `Program_01()` se spouští s nejnižší možnou prioritou v systému.
	To znamená, že ostatní úlohy, jako je komunikace, obsluha sítě Profinet atd., poběží s vyšší prioritou a budou `Program_01()` poměrně často přerušovat.
	Na druhou stranu `Program_01()` získá veškerý zbývající čas procesoru, což je obvykle více než 90&nbsp;%.
	Navíc je funkce `Program_01()` volána co nejrychleji (pokud je funkce téměř prázdná, interval volání je přibližně 200 - 400 ns).
	Tato koncepce umožňuje implementovat do funkce `Program_01()` jakékoli časově náročné výpočty bez obav z přetížení systému.
	Pro funkci `Program_01()` neexistuje chování v reálném čase.
	Doba cyklu v souboru `TGM.INI` musí být nastavena na nulu, jiná hodnota není povolena.
	Funkce `Program_01()` je volána vždy - nelze ji zakázat.
	
	``` ini
	[PLC_Configuration]
	Cycle_Time_Program_01=0
	```


V okně *System Timers* programu **Control Observer** se zobrazuje uplynulý čas různých úloh programu TGMotion.

![System timers dialog](../img/systemTimers.png){: style="width:60%;" }

Všimněte si, že maximální uplynulý čas úlohy `Program_01()` (označené jako PLC 1) zahrnuje také čas všech ostatních úloh, které přerušily úlohu `Program_01()`.
Podobně je tomu i u `Program_02()`, který může být přerušen `Programem_03()` a hlavní rutinou cyklu serva; a u `Program_03()`, přerušeného hlavní funkcí cyklu.
Maximální hodnotou je tedy součet uplynulého času funkce PLC a úloh s vyšší prioritou.
Jediné přesné měření času je uvedeno pro funkci `Program_04()`, protože ji nelze přerušit.

##Zařízení Profinet I/O
Zařízení Profinet I/O je ve výchozím nastavení zakázáno.
Použijte následující nastavení pro jeho povolení v souboru `TGM.INI`:

``` ini
[Profinet]
Enable=1
```

Každé zařízení Profinet je specifikováno svou MAC adresou, IP adresou a názvem zařízení.
Tyto atributy by měly být v rámci sítě Profinet jedinečné.
Zatímco adresa MAC je pevně daná, název a adresu IP lze nastavit pomocí nástroje pro uvedení do provozu (TIA Portal, Proneta atd.).
Ve výchozím nastavení je název Profinet TGMcontrolleru prázdný a IP adresa je nastavena na `0.0.0.0`.
Organizace modulů a slotů je dána konfiguračním souborem `GSDML`.
K dispozici je devět slotů: 1 × DAP, 4 vstupní sloty (velikost 32, 64, 128 nebo 256 bajtů) a 4 výstupní sloty (se stejnými velikostmi).
Každý slot má přiřazený parametr, který udává offset dat k paměti PLC DATA.
Tento parametr lze nastavit pomocí nástroje pro uvedení do provozu nebo pomocí programu Profinet PLC.

!!! warning "Změna IP"
	
	Změnou IP adresy zařízení Profinet se změní i adresa servisního portu X11.
	To znamená, že může dojít ke ztrátě TCP/UDP nebo jiné komunikace.
	Řešením je použití stejné adresy pro servisní protokoly a pro Profinet nebo použití portu FSP (X12) pro přímou komunikaci s PC, nezávisle na síti Profinet.
		
##Modbus TCP
Protokol Modbus TCP je ve výchozím nastavení vypnut.
Chcete-li jej povolit, použijte následující konfiguraci `TGM.INI`:
``` ini
[Modbus]
Povolit=1
```

Paměť PLC DATA se používá pro ukládání nebo příjem dat Modbus.
Číslo registru Modbus se používá jako offset v paměti.
Protože se čísla registrů počítají po jednom, je k dispozici násobič s výchozí hodnotou 4.
Také lze použít globální offset pro posun dat Modbus v paměti DATA (výchozí hodnota 0).
Doporučuje se, aby byl globální offset dělitelný číslem 4.
Konečná pozice dat Modbus se vypočítá podle následujícího vzorce:

`final_offset_to_data_memory = global_offset + (multiplier × modbus_register_number)`

Globální offset a násobitel lze nastavit také v souboru `TGM.INI`. Například:

``` ini
[Modbus]
Enable=1
Offset=16384
Multiplier=4
```

Podporovány jsou následující příkazy Modbus:

| příkaz  | číslo |
|----------|----------|
| READ_COILS            | 1  |
| READ_DISCRETE_INPUTS  | 2  |
| READ_HOLDING_REGISTERS| 3  |
| READ_INPUT_REGISTERS  | 4  |
| WRITE_SINGLE_REGISTER | 6  |
| WRITE_MULTIPLE_COILS  | 15 |
| WRITE_MULTIPLE_REGISTERS| 16 |

##Aktualizace firmwaru
Firmware lze snadno aktualizovat pomocí programu **Control Observer**.
Stačí vybrat správný soubor v editačním poli *Systémový soubor TGM* a použít tlačítko `[Write]`.
Soubor může být v počítači uložen pod libovolným názvem.
Při ukládání souboru na kartu SD použije TGMController správný název souboru.
Po přenosu souboru se zařízení automaticky restartuje.
Aktualizovaný firmware se uloží na kartu SD.
Soubor je možné uložit také přímo na SD kartu pomocí PC, v takovém případě musí být název `TGMC.fw` a soubor musí být v kořenovém adresáři karty.
Viz také další kapitola o obsahu SD karty.

![System files FW dialog](../img/systemFilesFW.png){: style="width:60%;" }

Může se stát, že se při zpětném čtení souborů ze zařízení se objeví chyba `Device Offline`, zejména při použití protokolu FSP.
Obvykle to znamená, že použitý ethernetový adaptér nemá dostatek vyrovnávacích paměti (doporučuje se 32 nebo více).
Protokol FSP je určen pro výkonné adaptéry, konkrétně pro adaptéry PCIe.

##Obsah karty SD
Na kartě SD musí být uloženy dva důležité a povinné soubory s pevnými názvy:

- `TGMC.fw` - soubor firmwaru (ve speciálním binárním formátu)
- `TGM.INI` - konfigurační textový soubor

Existují také možné další soubory:

- `TGM_PLC.bin` - virtuální PLC vytvořené uživatelem
- `Eeprom.bin` - binární konfigurační soubor pro PLC

Všechny čtyři výše uvedené soubory lze zapsat na kartu SD pomocí servisního programu **Control Observer** v počítači.
Tyto soubory lze také vyčíst zpět.
Protože karta SD používá standardní souborový systém FAT32, lze k ní také přímo přistupovat pomocí počítače a čtečky karet SD.
Po vložení upravené SD karty zpět do řídicí jednotky TGMcontroller je nutné ji restartovat pomocí programu **Control Observer** nebo sekvencí vypnutí/zapnutí.
Názvy těchto souborů jsou pevně dané.

## Safe boot
TGMcontroller se normálně zavádí z karty SD.
V případě vadné karty nebo poškozeného firmwaru je možné zařízení spustit z interní flash paměti.
Vyjměte SD kartu a restartujte systém.
Bez SD karty trvá spuštění dlouho.
Systém načte výchozí obsah souboru `INI` s IP adresou `192.168.1.188`.
Pokud je to možné, použijte port FSP, protože na straně počítače není nutné žádné nastavení.
Vložte novou SD kartu naformátovanou na FAT32 a nahrajte všechny potřebné soubory pomocí **Control Observer**: `TGM.INI` a poté firmware TGMcontroller.
Zařízení se automaticky restartuje a mělo by se správně spustit z SD karty.   

Soubor `TGM.INI` je uložen na kartě SD v přirozené textové podobě, takže v případě chybného nastavení je možné jej upravit přímo pomocí počítače.
Po vložení SD karty zpět do přístroje je nutné provést restart zařízení.










