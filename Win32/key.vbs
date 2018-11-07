Dim edition
edition = WScript.Arguments(0)
Set keys = CreateObject ("Scripting.Dictionary")

'Windows 10
keys.Add "58e97c99-f377-4ef1-81d5-4ad5522b5fd8", "TX9XD-98N7V-6WMQ6-BX7FG-H8Q99" 'Home
keys.Add "7b9e1751-a8da-4f75-9560-5fadfe3d8e38", "3KHY7-WNT83-DGQKR-F7HPR-844BM" 'Home N
keys.Add "cd918a57-a41b-4c82-8dce-1a538e221a83", "7HNRX-D7KGG-3K4RQ-4WPJ4-YTDFH" 'Home Single Language
keys.Add "a9107544-f4a0-4053-a96a-1479abdef912", "PVMJN-6DFY6-9CCP6-7BKTT-D3WVR" 'Home China
keys.Add "2de67392-b7a7-462a-b1ca-108dd189f588", "W269N-WFGWX-YVC9B-4J6C9-T83GX" 'Pro
keys.Add "a80b5abf-76ad-428b-b05d-a47d2dffeebf", "MH37W-N47XK-V7XM9-C7227-GCQG9" 'Pro N
keys.Add "3f1afc82-f8ac-4f6c-8005-1d233e606eee", "6TP4R-GNPTD-KYYHQ-7B7DP-J447Y" 'Pro Education
keys.Add "5300b18c-2e33-4dc2-8291-47ffcec746dd", "YVWGF-BXNMC-HTQYQ-CPQ99-66QFC" 'Pro Education N
keys.Add "82bbc092-bc50-4e16-8e18-b74fc486aec3", "NRG8B-VKK3Q-CXVCJ-9G2XF-6Q84J" 'Pro Workstation
keys.Add "4b1571d3-bafb-4b40-8087-a961be2caf65", "9FNHH-K3HBT-3W4TD-6383H-6XYWF" 'Pro Workstation N
keys.Add "e0c42288-980c-4788-a014-c080d2e1926e", "NW6C2-QMPVW-D7KKK-3GKT6-VCFB2" 'Education
keys.Add "3c102355-d027-42c6-ad23-2e7ef8a02585", "2WH4N-8QGBV-H22JP-CT43Q-MDWWJ" 'Education N
keys.Add "73111121-5638-40f6-bc11-f1d7b0d64300", "NPPR9-FWDCX-D2C8J-H872K-2YT43" 'Enterprise
keys.Add "e272e3e2-732f-4c65-a8f0-484747d0d947", "DPH2V-TTNVB-4X9Q3-TJR4H-KHJW4" 'Enterprise N
keys.Add "e0b2d383-d112-413f-8a80-97f373a5820c", "YYVX9-NTFWV-6MDM3-9PT4T-4M68B" 'Enterprise G
keys.Add "e38454fb-41a4-4f59-a5dc-25080e354730", "44RPN-FTY23-9VTTB-MP9BX-T84FV" 'Enterprise G N
keys.Add "7b51a46c-0c04-4e8f-9af4-8496cca90d5e", "WNMTR-4C88C-JK8YV-HQ7T2-76DF9" 'Enterprise 2015 LTSB
keys.Add "87b838b7-41b6-4590-8318-5797951d8529", "2F77B-TNFGY-69QQF-B8YKP-D69TJ" 'Enterprise 2015 LTSB N
keys.Add "2d5a5a60-3040-48bf-beb0-fcd770c20ce0", "DCPHK-NFMTC-H88MJ-PFHPY-QJ4BJ" 'Enterprise 2016 LTSB
keys.Add "9f776d83-7156-45b2-8a5c-359b9c9f22a3", "QFFDN-GRT3P-VKWWX-X7T3R-8B639" 'Enterprise 2016 LTSB N
keys.Add "32d2fab3-e4a8-42c2-923b-4bf4fd13e6ee", "M7XTQ-FN8P6-TTKYV-9D4CC-J462D" 'Enterprise LTSC 2019
keys.Add "7103a333-b8c8-49cc-93ce-d37c09687f92", "92NFX-8DJQP-P6BBQ-THF9C-7CG2H" 'Enterprise LTSC 2019 N
keys.Add "ec868e65-fadf-4759-b23e-93fe37f2cc29", "CPWHC-NT2C7-VYW78-DHDB2-PG3GK" 'Enterprise for Virtual Desktops
keys.Add "e4db50ea-bda1-4566-b047-0ca50abc6f07", "7NBT4-WGBQX-MP4H7-QXFF8-YP3KX" 'Remote Server
keys.Add "0df4f814-3f57-4b8b-9a9d-fddadcd69fac", "NBTWJ-3DR69-3C4V8-C26MC-GQ9M6" 'Lean

'Windows Server 2019
keys.Add "de32eafd-aaee-4662-9444-c1befb41bde2", "N69G4-B89J2-4G8F4-WWYCC-J464C" 'Standard
keys.Add "34e1ae55-27f8-4950-8877-7a03be5fb181", "WMDGN-G9PQG-XVVXX-R3X43-63DFG" 'Datacenter
keys.Add "034d3cbb-5d4b-4245-b3f8-f84571314078", "WVDHN-86M7X-466P6-VHXV7-YY726" 'Essentials
keys.Add "a99cc1f0-7719-4306-9645-294102fbff95", "FDNH6-VW9RW-BXPJ7-4XTYG-239TB" 'Azure Core
keys.Add "73e3957c-fc0c-400d-9184-5f7b6f2eb409", "N2KJX-J94YW-TQVFB-DG9YT-724CC" 'Standard ACor
keys.Add "90c362e5-0da1-4bfd-b53b-b87d309ade43", "6NMRW-2C8FM-D24W7-TQWMY-CWH2D" 'Datacenter ACor
keys.Add "8de8eb62-bbe0-40ac-ac17-f75595071ea3", "GRFBW-QNDC4-6QBHG-CCK3B-2PR88" 'ServerARM64

'Windows Server 2016
keys.Add "8c1c5410-9f39-4805-8c9d-63a07706358f", "WC2BQ-8NRM3-FDDYY-2BFGV-KHKQY" 'Standard
keys.Add "21c56779-b449-4d20-adfc-eece0e1ad74b", "CB7KF-BWN84-R7R2Y-793K2-8XDDG" 'Datacenter
keys.Add "2b5a1b0f-a5ab-4c54-ac2f-a6d94824a283", "JCKRF-N37P4-C2D82-9YXRT-4M63B" 'Essentials
keys.Add "7b4433f4-b1e7-4788-895a-c45378d38253", "QN4C6-GBJD2-FB422-GHWJK-GJG2R" 'Cloud Storage
keys.Add "3dbf341b-5f6c-4fa7-b936-699dce9e263f", "VP34G-4NPPG-79JTQ-864T4-R3MQX" 'Azure Core
keys.Add "61c5ef22-f14f-4553-a824-c4b31e84b100", "PTXN8-JFHJM-4WC78-MPCBR-9W4KR" 'Standard ACor
keys.Add "e49c08e7-da82-42f8-bde2-b570fbcae76c", "2HXDN-KRXHB-GPYC7-YCKFJ-7FVDG" 'Datacenter ACor
keys.Add "43d9af6e-5e86-4be8-a797-d072a046896c", "K9FYF-G6NCK-73M32-XMVPY-F9DRR" 'ServerARM64

'Windows 8.1
keys.Add "fe1c3238-432a-43a1-8e25-97e7d1ef10f3", "M9Q9P-WNJJT-6PXPY-DWX8H-6XWKK" 'Core
keys.Add "78558a64-dc19-43fe-a0d0-8075b2a370a3", "7B9N3-D94CG-YTVHR-QBPX3-RJP64" 'Core N
keys.Add "c72c6a1d-f252-4e7e-bdd1-3fca342acb35", "BB6NG-PQ82V-VRDPW-8XVD2-V8P66" 'Core Single Language
keys.Add "db78b74f-ef1c-4892-abfe-1e66b8231df6", "NCTT7-2RGK8-WMHRF-RY7YQ-JTXG3" 'Core China
keys.Add "ffee456a-cd87-4390-8e07-16146c672fd0", "XYTND-K6QKT-K2MRH-66RTM-43JKP" 'Core ARM
keys.Add "c06b6981-d7fd-4a35-b7b4-054742b7af67", "GCRJD-8NW9H-F2CDX-CCM8D-9D6T9" 'Pro
keys.Add "7476d79f-8e48-49b4-ab63-4d0b813a16e4", "HMCNV-VVBFX-7HMBH-CTY9B-B4FXY" 'Pro N
keys.Add "096ce63d-4fac-48a9-82a9-61ae9e800e5f", "789NJ-TQK6T-6XTH8-J39CJ-J8D3P" 'Pro with Media Center
keys.Add "81671aaf-79d1-4eb1-b004-8cbbe173afea", "MHF9N-XY6XB-WVXMC-BTDCT-MKKG7" 'Enterprise
keys.Add "113e705c-fa49-48a4-beea-7dd879b46b14", "TT4HM-HN7YT-62K67-RGRQJ-JFFXW" 'Enterprise N
keys.Add "0ab82d54-47f4-4acb-818c-cc5bf0ecb649", "NMMPB-38DD4-R2823-62W8D-VXKJB" 'Embedded Industry Pro
keys.Add "cd4e2d9f-5059-4a50-a92d-05d5bb1267c7", "FNFKF-PWTVT-9RC8H-32HB2-JB34X" 'Embedded Industry Enterprise
keys.Add "f7e88590-dfc7-4c78-bccb-6f3865b99d1a", "VHXM3-NR6FT-RY6RT-CK882-KW2CJ" 'Embedded Industry Automotive
keys.Add "e9942b32-2e55-4197-b0bd-5ff58cba8860", "3PY8R-QHNP9-W7XQD-G6DPH-3J2C9" 'with Bing
keys.Add "c6ddecd6-2354-4c19-909b-306a3058484e", "Q6HTR-N24GM-PMJFP-69CD8-2GXKR" 'with Bing N
keys.Add "b8f5e3a3-ed33-4608-81e1-37d6c9dcfd9c", "KF37N-VDV38-GRRTV-XH8X6-6F3BB" 'with Bing Single Language
keys.Add "ba998212-460a-44db-bfb5-71bf09d1c68b", "R962J-37N87-9VVK2-WJ74P-XTMHR" 'with Bing China
keys.Add "e58d87b5-8126-4580-80fb-861b22f79296", "MX3RK-9HNGX-K3QKC-6PJ3F-W8D7B" 'Pro for Students
keys.Add "cab491c7-a918-4f60-b502-dab75e334f40", "TNFGH-2R6PB-8XM3K-QYHX2-J4296" 'Pro for Students N

'Windows Server 2012 R2
keys.Add "b3ca044e-a358-4d68-9883-aaa2941aca99", "D2N9P-3P6X9-2R39C-7RTCD-MDVJX" 'Standard
keys.Add "00091344-1ea4-4f37-b789-01750ba6988c", "W3GGN-FT8W3-Y4M27-J84CP-Q3VJ9" 'Datacenter
keys.Add "21db6ba4-9a7b-4a14-9e29-64a60c59301d", "KNC87-3J2TX-XB4WP-VCPJV-M4FWM" 'Essentials
keys.Add "b743a2be-68d4-4dd3-af32-92425b7bb623", "3NPTF-33KPT-GGBPR-YX76B-39KDD" 'Cloud Storage

'Windows 8
keys.Add "c04ed6bf-55c8-4b47-9f8e-5a1f31ceee60", "BN3D2-R7TKB-3YPBD-8DRP2-27GG4" 'Core
keys.Add "197390a0-65f6-4a95-bdc4-55d58a3b0253", "8N2M2-HWPGY-7PGT9-HGDD8-GVGGY" 'Core N
keys.Add "8860fcd4-a77b-4a20-9045-a150ff11d609", "2WN2H-YGCQR-KFX6K-CD6TF-84YXQ" 'Core Single Language
keys.Add "9d5584a2-2d85-419a-982c-a00888bb9ddf", "4K36P-JN4VD-GDC6V-KDT89-DYFKP" 'Core China
keys.Add "af35d7b7-5035-4b63-8972-f0b747b9f4dc", "DXHJF-N9KQX-MFPVR-GHGQK-Y7RKV" 'Core ARM
keys.Add "a98bcd6d-5343-4603-8afe-5908e4611112", "NG4HW-VH26C-733KW-K6F98-J8CK4" 'Pro
keys.Add "ebf245c1-29a8-4daf-9cb1-38dfc608a8c8", "XCVCF-2NXM9-723PB-MHCB7-2RYQQ" 'Pro N
keys.Add "a00018a3-f20f-4632-bf7c-8daa5351c914", "GNBB8-YVD74-QJHX6-27H4K-8QHDG" 'Pro with Media Center
keys.Add "458e1bec-837a-45f6-b9d5-925ed5d299de", "32JNW-9KQ84-P47T8-D8GGY-CWCK7" 'Enterprise
keys.Add "e14997e7-800a-4cf7-ad10-de4b45b578db", "JMNMF-RHW7P-DMY6X-RF3DR-X2BQT" 'Enterprise N
keys.Add "10018baf-ce21-4060-80bd-47fe74ed4dab", "RYXVT-BNQG7-VD29F-DBMRY-HT73M" 'Embedded Industry Pro
keys.Add "18db1848-12e0-4167-b9d7-da7fcda507db", "NKB3R-R2F8T-3XCDP-7Q2KW-XWYQ2" 'Embedded Industry Enterprise

'Windows Server 2012
keys.Add "f0f5ec41-0d55-4732-af02-440a44a3cf0f", "XC9B7-NBPP2-83J2H-RHMBY-92BT4" 'Standard
keys.Add "d3643d60-0c42-412d-a7d6-52e6635327f6", "48HP8-DN98B-MYWDG-T2DCC-8W83P" 'Datacenter
keys.Add "7d5486c7-e120-4771-b7f1-7b56c6d3170c", "HM7DN-YVMH3-46JC3-XYTG7-CYQJJ" 'MultiPoint Standard
keys.Add "95fd1c83-7df5-494a-be8b-1300e1c9d1cd", "XNH6W-2V9GX-RGJ4K-Y8X6F-QGJ2G" 'MultiPoint Premium

'Windows 7
keys.Add "b92e9980-b9d5-4821-9c94-140f632f6312", "FJ82H-XT6CR-J8D7P-XQJJ2-GPDD4" 'Professional
keys.Add "54a09a0d-d57b-4c10-8b69-a842d6590ad5", "MRPKT-YTG23-K7D7T-X2JMM-QY7MG" 'Professional N
keys.Add "5a041529-fef8-4d07-b06f-b59b573b32d2", "W82YF-2Q76Y-63HXB-FGJG9-GF7QX" 'Professional E
keys.Add "ae2ee509-1b34-41c0-acb7-6d4650168915", "33PXH-7Y6KF-2VJC9-XBBR8-HVTHH" 'Enterprise
keys.Add "1cb6d605-11b3-4e14-bb30-da91c8e3983a", "YDRBP-3D83W-TY26F-D46B2-XCKRJ" 'Enterprise N
keys.Add "46bbed08-9c7b-48fc-a614-95250573f4ea", "C29WB-22CC8-VJ326-GHFJW-H9DH4" 'Enterprise E
keys.Add "db537896-376f-48ae-a492-53d0547773d0", "YBYF6-BHCR3-JPKRB-CDW7B-F9BK4" 'Embedded POSReady 7
keys.Add "e1a8296a-db37-44d1-8cce-7bc961d59c54", "XGY72-BRBBT-FF8MH-2GG8H-W7KCW" 'Embedded Standard
keys.Add "aa6dd3aa-c2b4-40e2-a544-a6bbb3f5c395", "73KQT-CD9G6-K7TQG-66MRP-CQ22C" 'Embedded ThinPC

'Windows Server 2008 R2
keys.Add "a78b8bd9-8017-4df5-b86a-09f756affa7c", "6TPJF-RBVHG-WBW2R-86QPH-6RTM4" 'Web
keys.Add "cda18cf3-c196-46ad-b289-60c072869994", "TT8MH-CG224-D3D7Q-498W2-9QCTX" 'HPC
keys.Add "68531fb9-5511-4989-97be-d11a0f55633f", "YC6KT-GKW9T-YTKYR-T4X34-R7VHC" 'Standard
keys.Add "7482e61b-c589-4b7f-8ecc-46d455ac3b87", "74YFP-3QFB3-KQT8W-PMXWJ-7M648" 'Datacenter
keys.Add "620e2b3d-09e7-42fd-802a-17a13652fe7a", "489J6-VHDMP-X63PK-3K798-CPX3Y" 'Enterprise
keys.Add "8a26851c-1c7e-48d3-a687-fbca9b9ac16b", "GT63C-RJFQ3-4GMB6-BRFB9-CB83V" 'Itanium
keys.Add "f772515c-0e87-48d5-a676-e6962c3e1195", "736RG-XDKJK-V34PF-BHK87-J6X3K" 'MultiPoint Server

'Office 2019
keys.Add "0bc88885-718c-491d-921f-6f214349e79c", "VQ9DP-NVHPH-T9HJC-J9PDT-KTQRG" 'Professional Plus C2R-P
keys.Add "fc7c4d0c-2e85-4bb9-afd4-01ed1476b5e9", "XM2V9-DN9HH-QB449-XDGKC-W2RMW" 'Project Professional C2R-P
keys.Add "500f6619-ef93-4b75-bcb4-82819998a3ca", "N2CG9-YD3YK-936X4-3WR82-Q3X4H" 'Visio Professional C2R-P
keys.Add "85dd8b5f-eaa4-4af3-a628-cce9e77c9a03", "NMMKJ-6RK4F-KMJVX-8D9MJ-6MWKP" 'Professional Plus
keys.Add "6912a74b-a5fb-401a-bfdb-2e3ab46f4b02", "6NWWJ-YQWMR-QKGCB-6TMB3-9D9HK" 'Standard
keys.Add "2ca2bf3f-949e-446a-82c7-e25a15ec78c4", "B4NPR-3FKK7-T2MBV-FRQ4W-PKD2B" 'Project Professional
keys.Add "1777f0e3-7392-4198-97ea-8ae4de6f6381", "C4F7P-NCP8C-6CQPT-MQHV9-JXD2M" 'Project Standard
keys.Add "5b5cf08f-b81a-431d-b080-3450d8620565", "9BGNQ-K37YR-RQHF2-38RQ3-7VCBB" 'Visio Professional
keys.Add "e06d7df3-aad0-419d-8dfb-0ac37e2bdf39", "7TQNQ-K3YQQ-3PFH7-CCPPM-X4VQ2" 'Visio Standard
keys.Add "9e9bceeb-e736-4f26-88de-763f87dcc485", "9N9PT-27V4Y-VJ2PD-YXFMF-YTFQT" 'Access
keys.Add "237854e9-79fc-4497-a0c1-a70969691c6b", "TMJWT-YYNMB-3BKTF-644FC-RVXBD" 'Excel
keys.Add "c8f8a301-19f5-4132-96ce-2de9d4adbd33", "7HD7K-N4PVK-BHBCQ-YWQRW-XW4VK" 'Outlook
keys.Add "3131fd61-5e4f-4308-8d6d-62be1987c92c", "RRNCX-C64HY-W2MM7-MCH9G-TJHMQ" 'PowerPoint
keys.Add "9d3e4cca-e172-46f1-a2f4-1d2107051444", "G2KWX-3NW6P-PY93R-JXK2T-C9Y9V" 'Publisher
keys.Add "734c6c6e-b0ba-4298-a891-671772b2bd1b", "NCJ33-JHBBY-HTK98-MYCV8-HMKHJ" 'Skype for Business
keys.Add "059834fe-a8ea-4bff-b67b-4d006b5447d3", "PBX3G-NWMT6-Q7XBW-PYJGG-WXD33" 'Word

'Office 2016
keys.Add "829b8110-0e6f-4349-bca4-42803577788d", "WGT24-HCNMF-FQ7XH-6M8K7-DRTW9" 'Project Professional C2R-P
keys.Add "cbbaca45-556a-4416-ad03-bda598eaa7c8", "D8NRQ-JTYM3-7J2DX-646CT-6836M" 'Project Standard C2R-P
keys.Add "b234abe3-0857-4f9c-b05a-4dc314f85557", "69WXN-MBYV6-22PQG-3WGHK-RM6XC" 'Visio Professional C2R-P
keys.Add "361fe620-64f4-41b5-ba77-84f8e079b1f7", "NY48V-PPYYH-3F4PX-XJRKJ-W4423" 'Visio Standard C2R-P
keys.Add "e914ea6e-a5fa-4439-a394-a9bb3293ca09", "DMTCJ-KNRKX-26982-JYCKT-P7KB6" 'MondoR
keys.Add "9caabccb-61b1-4b4b-8bec-d10a3c3ac2ce", "HFTND-W9MK4-8B7MJ-B6C4G-XQBR2" 'Mondo
keys.Add "d450596f-894d-49e0-966a-fd39ed4c4c64", "XQNVK-8JYDB-WJ9W3-YJ8YR-WFG99" 'Professional Plus
keys.Add "dedfa23d-6ed1-45a6-85dc-63cae0546de6", "JNRGM-WHDWX-FJJG3-K47QV-DRTFM" 'Standard
keys.Add "4f414197-0fc2-4c01-b68a-86cbb9ac254c", "YG9NW-3K39V-2T3HJ-93F3Q-G83KT" 'Project Professional
keys.Add "da7ddabc-3fbe-4447-9e01-6ab7440b4cd4", "GNFHQ-F6YQM-KQDGJ-327XX-KQBVC" 'Project Standard
keys.Add "6bf301c1-b94a-43e9-ba31-d494598c47fb", "PD3PC-RHNGV-FXJ29-8JK7D-RJRJK" 'Visio Professional
keys.Add "aa2a7821-1827-4c2c-8f1d-4513a34dda97", "7WHWN-4T7MP-G96JF-G33KR-W8GF4" 'Visio Standard
keys.Add "67c0fc0c-deba-401b-bf8b-9c8ad8395804", "GNH9Y-D2J4T-FJHGG-QRVH7-QPFDW" 'Access
keys.Add "c3e65d36-141f-4d2f-a303-a842ee756a29", "9C2PK-NWTVB-JMPW8-BFT28-7FTBF" 'Excel
keys.Add "d8cace59-33d2-4ac7-9b1b-9b72339c51c8", "DR92N-9HTF2-97XKM-XW2WJ-XW3J6" 'OneNote
keys.Add "ec9d9265-9d1e-4ed0-838a-cdc20f2551a1", "R69KK-NTPKF-7M3Q4-QYBHW-6MT9B" 'Outlook
keys.Add "d70b1bba-b893-4544-96e2-b7a318091c33", "J7MQP-HNJ4Y-WJ7YM-PFYGF-BY6C6" 'Powerpoint
keys.Add "041a06cb-c5b8-4772-809f-416d03d16654", "F47MM-N3XJP-TQXJ9-BP99D-8K837" 'Publisher
keys.Add "83e04ee1-fa8d-436d-8994-d31a862cab77", "869NQ-FJ69K-466HW-QYCP2-DDBV6" 'Skype for Business
keys.Add "bb11badf-d8aa-470e-9311-20eaf80fe5cc", "WXY84-JN2Q9-RBCCQ-3Q3J3-3PFJ6" 'Word

'Office 2013
keys.Add "dc981c6b-fc8e-420f-aa43-f8f33e5c0923", "42QTK-RN8M7-J3C4G-BBGYM-88CYV" 'Mondo
keys.Add "b322da9c-a2e2-4058-9e4e-f59a6970bd69", "YC7DK-G2NP3-2QQC3-J6H88-GVGXT" 'Professional Plus
keys.Add "b13afb38-cd79-4ae5-9f7f-eed058d750ca", "KBKQT-2NMXY-JJWGP-M62JB-92CD4" 'Standard
keys.Add "4a5d124a-e620-44ba-b6ff-658961b33b9a", "FN8TT-7WMH6-2D4X9-M337T-2342K" 'Project Professional
keys.Add "427a28d1-d17c-4abf-b717-32c780ba6f07", "6NTH3-CW976-3G3Y2-JK3TX-8QHTT" 'Project Standard
keys.Add "e13ac10e-75d0-4aff-a0cd-764982cf541c", "C2FG9-N6J68-H8BTJ-BW3QX-RM3B3" 'Visio Professional
keys.Add "ac4efaf0-f81f-4f61-bdf7-ea32b02ab117", "J484Y-4NKBF-W2HMG-DBMJC-PGWR7" 'Visio Standard
keys.Add "6ee7622c-18d8-4005-9fb7-92db644a279b", "NG2JY-H4JBT-HQXYP-78QH9-4JM2D" 'Access
keys.Add "f7461d52-7c2b-43b2-8744-ea958e0bd09a", "VGPNG-Y7HQW-9RHP7-TKPV3-BG7GB" 'Excel
keys.Add "fb4875ec-0c6b-450f-b82b-ab57d8d1677f", "H7R7V-WPNXQ-WCYYC-76BGV-VT7GH" 'Groove
keys.Add "a30b8040-d68a-423f-b0b5-9ce292ea5a8f", "DKT8B-N7VXH-D963P-Q4PHY-F8894" 'InfoPath
keys.Add "1b9f11e3-c85c-4e1b-bb29-879ad2c909e3", "2MG3G-3BNTT-3MFW9-KDQW3-TCK7R" 'Lync
keys.Add "efe1f3e6-aea2-4144-a208-32aa872b6545", "TGN6P-8MMBC-37P2F-XHXXK-P34VW" 'OneNote
keys.Add "771c3afa-50c5-443f-b151-ff2546d863a0", "QPN8Q-BJBTJ-334K3-93TGY-2PMBT" 'Outlook
keys.Add "8c762649-97d1-4953-ad27-b7e2c25b972e", "4NT99-8RJFH-Q2VDH-KYG2C-4RD4F" 'Powerpoint
keys.Add "00c79ff1-6850-443d-bf61-71cde0de305f", "PN2WF-29XG2-T9HJ7-JQPJR-FCXK4" 'Publisher
keys.Add "d9f5b1c6-5386-495a-88f9-9ad6b41ac9b3", "6Q7VD-NX8JD-WJ2VH-88V73-4GBJ7" 'Word

'Office 2010
keys.Add "09ed9640-f020-400a-acd8-d7d867dfd9c2", "YBJTT-JG6MD-V9Q7P-DBKXJ-38W9R" 'Mondo
keys.Add "ef3d4e49-a53d-4d81-a2b1-2ca6c2556b2c", "7TC2V-WXF6P-TD7RT-BQRXR-B8K32" 'Mondo2
keys.Add "6f327760-8c5c-417c-9b61-836a98287e0c", "VYBBJ-TRJPB-QFQRF-QFT4D-H3GVB" 'Professional Plus
keys.Add "9da2a678-fb6b-4e67-ab84-60dd6a9c819a", "V7QKV-4XVVR-XYV4D-F7DFM-8R6BM" 'Standard
keys.Add "df133ff7-bf14-4f95-afe3-7b48e7e331ef", "YGX6F-PGV49-PGW3J-9BTGG-VHKC6" 'Project Professional
keys.Add "5dc7bf61-5ec9-4996-9ccb-df806a2d0efe", "4HP3K-88W3F-W2K3D-6677X-F9PGB" 'Project Standard
keys.Add "92236105-bb67-494f-94c7-7f7a607929bd", "D9DWC-HPYVV-JGF4P-BTWQB-WX8BJ" 'Visio Premium
keys.Add "e558389c-83c3-4b29-adfe-5e4d7f46c358", "7MCW8-VRQVK-G677T-PDJCM-Q8TCP" 'Visio Professional
keys.Add "9ed833ff-4f92-4f36-b370-8683a4f13275", "767HD-QGMWX-8QTDB-9G3R2-KHFGJ" 'Visio Standard
keys.Add "8ce7e872-188c-4b98-9d90-f8f90b7aad02", "V7Y44-9T38C-R2VJK-666HK-T7DDX" 'Access
keys.Add "cee5d470-6e3b-4fcc-8c2b-d17428568a9f", "H62QG-HXVKF-PP4HP-66KMR-CW9BM" 'Excel
keys.Add "8947d0b8-c33b-43e1-8c56-9b674c052832", "QYYW6-QP4CB-MBV6G-HYMCJ-4T3J4" 'Groove (SharePoint Workspace)
keys.Add "ca6b6639-4ad6-40ae-a575-14dee07f6430", "K96W8-67RPQ-62T9Y-J8FQJ-BT37T" 'InfoPath
keys.Add "ab586f5c-5256-4632-962f-fefd8b49e6f4", "Q4Y4M-RHWJM-PY37F-MTKWH-D3XHX" 'OneNote
keys.Add "ecb7c192-73ab-4ded-acf4-2399b095d0cc", "7YDC2-CWM8M-RRTJC-8MDVC-X3DWQ" 'Outlook
keys.Add "45593b1d-dfb1-4e91-bbfb-2d5d0ce2227a", "RC8FX-88JRY-3PF7C-X8P67-P4VTT" 'Powerpoint
keys.Add "b50c4f75-599b-43e8-8dcd-1081a7967241", "BFK7F-9MYHM-V68C7-DRQ66-83YTP" 'Publisher
keys.Add "2d0882e7-a4e7-423b-8ccc-70d91e0158b1", "HVHB3-C6FV7-KQX9W-YQG79-CRY7T" 'Word
keys.Add "ea509e87-07a1-4a45-9edc-eba5a39f36af", "D6QFG-VBYP2-XQHM7-J97RH-VVRCK" 'Home and Business

if keys.Exists(edition) then
WScript.Echo keys.Item(edition)
End If
