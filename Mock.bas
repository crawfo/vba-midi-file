Attribute VB_Name = "Mock"
Option Explicit

Private Function ToArrayFromVariantArray(variantArray As Variant) As Byte()
    Dim byteArray() As Byte
    Dim i As Long
    Dim upperBound As Long
    upperBound = UBound(variantArray)
    ReDim byteArray(upperBound)
    For i = 0 To upperBound
        byteArray(i) = variantArray(i)
    Next i
    ToArrayFromVariantArray = byteArray
End Function

Public Function GetTestHeaderChunkBytes() As Byte
    GetTestHeaderChunkBytes = Array(77, 84, 104, 100)
    
End Function

Public Function GetTestChannelEvent() As ChannelEvent
    Set GetTestChannelEvent = Factory.CreateNewChannelEvent(isRunStatus:=False, _
                                                            deltaTime:=50, _
                                                            absoluteTime:=100, _
                                                            midiStatus:=&H8, _
                                                            midiChannel:=&H3, _
                                                            eventData1:=64, _
                                                            eventData2:=127)
End Function

Public Function GetTestRunningStatus() As ChannelEvent
    Set GetTestRunningStatus = Factory.CreateNewChannelEvent(isRunStatus:=True, _
                                                             deltaTime:=50, _
                                                             absoluteTime:=100, _
                                                             midiStatus:=&H8, _
                                                             midiChannel:=&H3, _
                                                             eventData1:=64, _
                                                             eventData2:=127)
End Function

Public Function GetTestMetaEvent() As MetaEvent
    Dim evtData As Collection

    Set evtData = New Collection
    evtData.Add 15
    Set GetTestMetaEvent = Factory.CreateNewMetaEvent(deltaTime:=50, _
                                                      absoluteTime:=100, _
                                                      midiMetaType:=32, _
                                                      eventData:=evtData)
                                             
End Function

Public Function GetTestSystemExclusiveEvent() As SystemExclusiveEvent
    Dim evtData As Collection

    Set evtData = New Collection
    evtData.Add 10
    evtData.Add &HF7
    Set GetTestSystemExclusiveEvent = _
    Factory.CreateNewSystemExclusiveEvent(deltaTime:=50, _
                                          absoluteTime:=100, _
                                          midiStatus:=&HF0, _
                                          eventData:=evtData, _
                                          systemExType:=SystemExclusiveType.NORMAL)
End Function

Public Function GetTestTrack() As Byte()
    Dim trkChunkBytes() As Variant
    trkChunkBytes = _
    Array(77, 84, 114, 107, 0, 0, 1, 62, 0, 255, 33, 1, 0, 0, 255, 3, 11, 80, 49, _
          45, 68, 105, 115, 116, 76, 101, 97, 100, 0, 176, 0, 0, 0, 32, 64, 0, _
          192, 84, 0, 176, 7, 127, 129, 81, 91, 127, 1, 93, 127, 141, 202, 46, _
          144, 79, 114, 135, 64, 79, 0, 0, 78, 117, 133, 80, 74, 117, 129, 112, _
          78, 0, 133, 80, 78, 107, 129, 108, 74, 0, 129, 112, 78, 0, 4, 79, 94, 129 _
          , 108, 79, 0, 4, 78, 94, 129, 108, 78, 0, 4, 79, 94, 129, 108, 79, 0, 4, _
          76, 108, 137, 48, 76, 0, 0, 83, 117, 131, 96, 83, 0, 0, 79, 111, 129, 108, _
          79, 0, 4, 78, 114, 130, 100, 78, 0, 4, 76, 103, 130, 100, 76, 0, 4, 78, _
          119, 129, 108, 78, 0, 4, 76, 110, 130, 100, 76, 0, 4, 78, 97, 129, 108, _
          78, 0, 4, 71, 103, 130, 100, 71, 0, 4, 72, 108, 150, 60, 72, 0, 4, 74, 115, _
          135, 64, 74, 0, 0, 67, 108, 157, 124, 67, 0, 4, 79, 114, 135, 64, 79, 0, _
          0, 78, 117, 133, 80, 74, 117, 129, 112, 78, 0, 133, 80, 78, 107, 129, _
          108, 74, 0, 129, 112, 78, 0, 4, 79, 94, 129, 108, 79, 0, 4, 78, 94, 129, _
          108, 78, 0, 4, 79, 94, 129, 108, 79, 0, 4, 76, 108, 137, 48, 76, 0, 0, 83, _
          117, 131, 96, 83, 0, 0, 79, 111, 129, 108, 79, 0, 4, 78, 114, 130, 100, 78, _
          0, 4, 76, 103, 130, 100, 76, 0, 4, 78, 119, 129, 108, 78, 0, 4, 76, 110, 130, _
          100, 76, 0, 4, 78, 97, 129, 108, 78, 0, 4, 71, 103, 130, 100, 71, 0, 4, 72, 108, _
          150, 60, 72, 0, 4, 76, 115, 135, 64, 76, 0, 0, 74, 108, 157, 124, 74, 0, _
          0, 255, 47, 0)
    
    GetTestTrack = ListUtils.ToByteArray(trkChunkBytes)
End Function

Public Function GetTestFileWithWrongTrackLength() As Byte()
    Dim fileBytesV As Variant
    Dim FileBytes() As Byte
    Dim b As Variant
    Dim i As Long
    
    fileBytesV = Array(77, 84, 104, 100, 0, 0, 0, 6, 0, 1, 0, 1, 1, 224, 77, 84, 114, 107, 0, 0, _
                       0, 12, 0, 146, 60, 68, 0, 64, 65, 0, 67, 62, 0, 255, 47, 0)
    ReDim FileBytes(UBound(fileBytesV))
    For Each b In fileBytesV
        FileBytes(i) = b
        i = i + 1
    Next b
    GetTestFileWithWrongTrackLength = FileBytes
End Function

Public Function GetTestFile() As Byte()
    Dim tf(471) As Byte
    tf(0) = 77
    tf(1) = 84
    tf(2) = 104
    tf(3) = 100
    tf(4) = 0
    tf(5) = 0
    tf(6) = 0
    tf(7) = 6
    tf(8) = 0
    tf(9) = 1
    tf(10) = 0
    tf(11) = 2
    tf(12) = 1
    tf(13) = 224
    tf(14) = 77
    tf(15) = 84
    tf(16) = 114
    tf(17) = 107
    tf(18) = 0
    tf(19) = 0
    tf(20) = 0
    tf(21) = 124
    tf(22) = 0
    tf(23) = 255
    tf(24) = 3
    tf(25) = 8
    tf(26) = 117
    tf(27) = 110
    tf(28) = 116
    tf(29) = 105
    tf(30) = 116
    tf(31) = 108
    tf(32) = 101
    tf(33) = 100
    tf(34) = 0
    tf(35) = 255
    tf(36) = 33
    tf(37) = 1
    tf(38) = 0
    tf(39) = 0
    tf(40) = 240
    tf(41) = 5
    tf(42) = 126
    tf(43) = 127
    tf(44) = 9
    tf(45) = 1
    tf(46) = 247
    tf(47) = 0
    tf(48) = 240
    tf(49) = 8
    tf(50) = 67
    tf(51) = 16
    tf(52) = 76
    tf(53) = 0
    tf(54) = 0
    tf(55) = 126
    tf(56) = 0
    tf(57) = 247
    tf(58) = 0
    tf(59) = 240
    tf(60) = 8
    tf(61) = 67
    tf(62) = 16
    tf(63) = 76
    tf(64) = 2
    tf(65) = 1
    tf(66) = 12
    tf(67) = 109
    tf(68) = 247
    tf(69) = 0
    tf(70) = 240
    tf(71) = 9
    tf(72) = 67
    tf(73) = 16
    tf(74) = 76
    tf(75) = 2
    tf(76) = 1
    tf(77) = 32
    tf(78) = 66
    tf(79) = 0
    tf(80) = 247
    tf(81) = 0
    tf(82) = 240
    tf(83) = 8
    tf(84) = 67
    tf(85) = 16
    tf(86) = 76
    tf(87) = 2
    tf(88) = 1
    tf(89) = 44
    tf(90) = 127
    tf(91) = 247
    tf(92) = 0
    tf(93) = 255
    tf(94) = 33
    tf(95) = 1
    tf(96) = 0
    tf(97) = 0
    tf(98) = 255
    tf(99) = 88
    tf(100) = 4
    tf(101) = 4
    tf(102) = 2
    tf(103) = 24
    tf(104) = 8
    tf(105) = 0
    tf(106) = 255
    tf(107) = 89
    tf(108) = 2
    tf(109) = 0
    tf(110) = 0
    tf(111) = 0
    tf(112) = 255
    tf(113) = 81
    tf(114) = 3
    tf(115) = 9
    tf(116) = 39
    tf(117) = 192
    tf(118) = 131
    tf(119) = 67
    tf(120) = 255
    tf(121) = 6
    tf(122) = 10
    tf(123) = 88
    tf(124) = 71
    tf(125) = 101
    tf(126) = 100
    tf(127) = 105
    tf(128) = 116
    tf(129) = 32
    tf(130) = 69
    tf(131) = 110
    tf(132) = 100
    tf(133) = 129
    tf(134) = 161
    tf(135) = 61
    tf(136) = 255
    tf(137) = 81
    tf(138) = 3
    tf(139) = 6
    tf(140) = 126
    tf(141) = 60
    tf(142) = 0
    tf(143) = 255
    tf(144) = 47
    tf(145) = 0
    tf(146) = 77
    tf(147) = 84
    tf(148) = 114
    tf(149) = 107
    tf(150) = 0
    tf(151) = 0
    tf(152) = 1
    tf(153) = 62
    tf(154) = 0
    tf(155) = 255
    tf(156) = 33
    tf(157) = 1
    tf(158) = 0
    tf(159) = 0
    tf(160) = 255
    tf(161) = 3
    tf(162) = 11
    tf(163) = 80
    tf(164) = 49
    tf(165) = 45
    tf(166) = 68
    tf(167) = 105
    tf(168) = 115
    tf(169) = 116
    tf(170) = 76
    tf(171) = 101
    tf(172) = 97
    tf(173) = 100
    tf(174) = 0
    tf(175) = 176
    tf(176) = 0
    tf(177) = 0
    tf(178) = 0
    tf(179) = 32
    tf(180) = 64
    tf(181) = 0
    tf(182) = 192
    tf(183) = 84
    tf(184) = 0
    tf(185) = 176
    tf(186) = 7
    tf(187) = 127
    tf(188) = 129
    tf(189) = 81
    tf(190) = 91
    tf(191) = 127
    tf(192) = 1
    tf(193) = 93
    tf(194) = 127
    tf(195) = 141
    tf(196) = 202
    tf(197) = 46
    tf(198) = 144
    tf(199) = 79
    tf(200) = 114
    tf(201) = 135
    tf(202) = 64
    tf(203) = 79
    tf(204) = 0
    tf(205) = 0
    tf(206) = 78
    tf(207) = 117
    tf(208) = 133
    tf(209) = 80
    tf(210) = 74
    tf(211) = 117
    tf(212) = 129
    tf(213) = 112
    tf(214) = 78
    tf(215) = 0
    tf(216) = 133
    tf(217) = 80
    tf(218) = 78
    tf(219) = 107
    tf(220) = 129
    tf(221) = 108
    tf(222) = 74
    tf(223) = 0
    tf(224) = 129
    tf(225) = 112
    tf(226) = 78
    tf(227) = 0
    tf(228) = 4
    tf(229) = 79
    tf(230) = 94
    tf(231) = 129
    tf(232) = 108
    tf(233) = 79
    tf(234) = 0
    tf(235) = 4
    tf(236) = 78
    tf(237) = 94
    tf(238) = 129
    tf(239) = 108
    tf(240) = 78
    tf(241) = 0
    tf(242) = 4
    tf(243) = 79
    tf(244) = 94
    tf(245) = 129
    tf(246) = 108
    tf(247) = 79
    tf(248) = 0
    tf(249) = 4
    tf(250) = 76
    tf(251) = 108
    tf(252) = 137
    tf(253) = 48
    tf(254) = 76
    tf(255) = 0
    tf(256) = 0
    tf(257) = 83
    tf(258) = 117
    tf(259) = 131
    tf(260) = 96
    tf(261) = 83
    tf(262) = 0
    tf(263) = 0
    tf(264) = 79
    tf(265) = 111
    tf(266) = 129
    tf(267) = 108
    tf(268) = 79
    tf(269) = 0
    tf(270) = 4
    tf(271) = 78
    tf(272) = 114
    tf(273) = 130
    tf(274) = 100
    tf(275) = 78
    tf(276) = 0
    tf(277) = 4
    tf(278) = 76
    tf(279) = 103
    tf(280) = 130
    tf(281) = 100
    tf(282) = 76
    tf(283) = 0
    tf(284) = 4
    tf(285) = 78
    tf(286) = 119
    tf(287) = 129
    tf(288) = 108
    tf(289) = 78
    tf(290) = 0
    tf(291) = 4
    tf(292) = 76
    tf(293) = 110
    tf(294) = 130
    tf(295) = 100
    tf(296) = 76
    tf(297) = 0
    tf(298) = 4
    tf(299) = 78
    tf(300) = 97
    tf(301) = 129
    tf(302) = 108
    tf(303) = 78
    tf(304) = 0
    tf(305) = 4
    tf(306) = 71
    tf(307) = 103
    tf(308) = 130
    tf(309) = 100
    tf(310) = 71
    tf(311) = 0
    tf(312) = 4
    tf(313) = 72
    tf(314) = 108
    tf(315) = 150
    tf(316) = 60
    tf(317) = 72
    tf(318) = 0
    tf(319) = 4
    tf(320) = 74
    tf(321) = 115
    tf(322) = 135
    tf(323) = 64
    tf(324) = 74
    tf(325) = 0
    tf(326) = 0
    tf(327) = 67
    tf(328) = 108
    tf(329) = 157
    tf(330) = 124
    tf(331) = 67
    tf(332) = 0
    tf(333) = 4
    tf(334) = 79
    tf(335) = 114
    tf(336) = 135
    tf(337) = 64
    tf(338) = 79
    tf(339) = 0
    tf(340) = 0
    tf(341) = 78
    tf(342) = 117
    tf(343) = 133
    tf(344) = 80
    tf(345) = 74
    tf(346) = 117
    tf(347) = 129
    tf(348) = 112
    tf(349) = 78
    tf(350) = 0
    tf(351) = 133
    tf(352) = 80
    tf(353) = 78
    tf(354) = 107
    tf(355) = 129
    tf(356) = 108
    tf(357) = 74
    tf(358) = 0
    tf(359) = 129
    tf(360) = 112
    tf(361) = 78
    tf(362) = 0
    tf(363) = 4
    tf(364) = 79
    tf(365) = 94
    tf(366) = 129
    tf(367) = 108
    tf(368) = 79
    tf(369) = 0
    tf(370) = 4
    tf(371) = 78
    tf(372) = 94
    tf(373) = 129
    tf(374) = 108
    tf(375) = 78
    tf(376) = 0
    tf(377) = 4
    tf(378) = 79
    tf(379) = 94
    tf(380) = 129
    tf(381) = 108
    tf(382) = 79
    tf(383) = 0
    tf(384) = 4
    tf(385) = 76
    tf(386) = 108
    tf(387) = 137
    tf(388) = 48
    tf(389) = 76
    tf(390) = 0
    tf(391) = 0
    tf(392) = 83
    tf(393) = 117
    tf(394) = 131
    tf(395) = 96
    tf(396) = 83
    tf(397) = 0
    tf(398) = 0
    tf(399) = 79
    tf(400) = 111
    tf(401) = 129
    tf(402) = 108
    tf(403) = 79
    tf(404) = 0
    tf(405) = 4
    tf(406) = 78
    tf(407) = 114
    tf(408) = 130
    tf(409) = 100
    tf(410) = 78
    tf(411) = 0
    tf(412) = 4
    tf(413) = 76
    tf(414) = 103
    tf(415) = 130
    tf(416) = 100
    tf(417) = 76
    tf(418) = 0
    tf(419) = 4
    tf(420) = 78
    tf(421) = 119
    tf(422) = 129
    tf(423) = 108
    tf(424) = 78
    tf(425) = 0
    tf(426) = 4
    tf(427) = 76
    tf(428) = 110
    tf(429) = 130
    tf(430) = 100
    tf(431) = 76
    tf(432) = 0
    tf(433) = 4
    tf(434) = 78
    tf(435) = 97
    tf(436) = 129
    tf(437) = 108
    tf(438) = 78
    tf(439) = 0
    tf(440) = 4
    tf(441) = 71
    tf(442) = 103
    tf(443) = 130
    tf(444) = 100
    tf(445) = 71
    tf(446) = 0
    tf(447) = 4
    tf(448) = 72
    tf(449) = 108
    tf(450) = 150
    tf(451) = 60
    tf(452) = 72
    tf(453) = 0
    tf(454) = 4
    tf(455) = 76
    tf(456) = 115
    tf(457) = 135
    tf(458) = 64
    tf(459) = 76
    tf(460) = 0
    tf(461) = 0
    tf(462) = 74
    tf(463) = 108
    tf(464) = 157
    tf(465) = 124
    tf(466) = 74
    tf(467) = 0
    tf(468) = 0
    tf(469) = 255
    tf(470) = 47
    tf(471) = 0

    GetTestFile = tf
End Function


