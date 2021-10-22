DELETE FROM [Ewcs].[dbo].[ForkliftTruckTransports]

INSERT INTO [Ewcs].[dbo].[ForkliftTruckTransports] ( 
ForkliftTruckTransportId, 
EnumForkliftTruckTransportStatus, 
AssignedSourceLocationId, 
AssignedDestinationLocationId, 
AssignedCarrierId, 
AssignedUnitId, 
AssignedCarrierIsVirtual, 
AllowDropWithoutScan )
VALUES
( 'TR-DEMO-P001', 'Idle', 'MAP.1', 'FIL.1', 'P2000', 'NFPM1_20090430-100A', 0, 1 ),
( 'TR-DEMO-P002', 'Idle', 'FIL.2', 'AT48.30.02', 'P2001', 'NFPM1_20090430-100A', 0, 1 ),
( 'TR-DEMO-P003', 'Idle', 'AT48.30.02', 'AT48.11.03', 'P2001', 'NFPM1_20090430-100A', 0, 1 ),
( 'TR-DEMO-P004', 'Idle', 'AT48.11.03', 'HU.1', 'P2001', 'NFPM1_20090430-100A', 0, 1 ),
( 'TR-DEMO-P005', 'Idle', 'HU.2', 'FC3.2', 'P2011', 'NFPM1_20090430-100A', 0, 1 ),
( 'TR-DEMO-P006', 'Idle', 'FC3.2', 'FC3.1', 'P2011', 'NFPM1_20090430-100A', 0, 1 ),
( 'TR-DEMO-P007', 'Idle', 'FC3.1', 'FC4.1', 'P2011', 'NFPM1_20090430-100A', 0, 1 ),
( 'TR-DEMO-P008', 'Idle', 'FC4.1', 'FIL.1', 'P2011', 'NFPM1_20090430-100A', 0, 1 ),
( 'TR-DEMO-R001', 'Idle', 'RAT.2', 'RDL.1', 'R1011', NULL, 0, 1 ),
( 'TR-DEMO-R002', 'Idle', 'RAT.2', 'RDL.2', 'R1012', NULL, 0, 1 ),
( 'TR-DEMO-R003', 'Idle', 'RAT.2', 'RDL.3', 'R1013', NULL, 0, 1 ),
( 'TR-DEMO-R004', 'Idle', 'RDL.2', 'RAT.1', 'R1012', NULL, 0, 1 ),
( 'TR-DEMO-R005', 'Idle', 'RDL.3', 'RAT.1', 'R1013', NULL, 0, 1 ),
( 'TR-DEMO-R006', 'Idle', 'RDL.1', 'RMAP.2', 'R1011', NULL, 0, 1 ),
( 'TR-DEMO-R007', 'Idle', 'RMAP.2', 'RMAP.1', 'R1011', NULL, 0, 1 ),
( 'TR-DEMO-R008', 'Idle', 'RMAP.1', 'RAT.1', 'R1011', NULL, 0, 1 )


DELETE FROM [Ewcs].[dbo].[ForkliftTruckTransportInLocationGroups]

INSERT INTO [Ewcs].[dbo].[ForkliftTruckTransportInLocationGroups]( 
ForkliftTruckTransportId, 
LocationGroupId )
 VALUES
( 'TR-DEMO-P001', 'MAP'),
( 'TR-DEMO-P002', 'Atelier48'),
( 'TR-DEMO-P003', 'Atelier48'),
( 'TR-DEMO-P004', 'Atelier48'),
( 'TR-DEMO-P005', 'FC3'),
( 'TR-DEMO-P006', 'FC3'),
( 'TR-DEMO-P007', 'FC4'),
( 'TR-DEMO-P008', 'MAP'),
( 'TR-DEMO-R001', 'KCP5'),
( 'TR-DEMO-R002', 'WIP2'),
( 'TR-DEMO-R003', 'FC4'),
( 'TR-DEMO-R004', 'WIP2'),
( 'TR-DEMO-R005', 'FC4'),
( 'TR-DEMO-R006', 'KCP5'),
( 'TR-DEMO-R007', 'MAP'),
( 'TR-DEMO-R008', 'MAP')


