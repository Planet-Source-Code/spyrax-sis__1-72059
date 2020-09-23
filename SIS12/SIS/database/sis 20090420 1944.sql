-- MySQL Administrator dump 1.4
--
-- ------------------------------------------------------
-- Server version	4.1.22-community-nt


/*!40101 SET @OLD_CHARACTER_SET_CLIENT=@@CHARACTER_SET_CLIENT */;
/*!40101 SET @OLD_CHARACTER_SET_RESULTS=@@CHARACTER_SET_RESULTS */;
/*!40101 SET @OLD_COLLATION_CONNECTION=@@COLLATION_CONNECTION */;
/*!40101 SET NAMES utf8 */;

/*!40014 SET @OLD_UNIQUE_CHECKS=@@UNIQUE_CHECKS, UNIQUE_CHECKS=0 */;
/*!40014 SET @OLD_FOREIGN_KEY_CHECKS=@@FOREIGN_KEY_CHECKS, FOREIGN_KEY_CHECKS=0 */;
/*!40101 SET @OLD_SQL_MODE=@@SQL_MODE, SQL_MODE='NO_AUTO_VALUE_ON_ZERO' */;


--
-- Create schema sis
--

CREATE DATABASE IF NOT EXISTS sis;
USE sis;

--
-- Definition of table `tblcourse`
--

DROP TABLE IF EXISTS `tblcourse`;
CREATE TABLE `tblcourse` (
  `course` text NOT NULL
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

--
-- Dumping data for table `tblcourse`
--

/*!40000 ALTER TABLE `tblcourse` DISABLE KEYS */;
INSERT INTO `tblcourse` (`course`) VALUES 
 ('BSCS'),
 ('BSIT'),
 ('BSCOE');
/*!40000 ALTER TABLE `tblcourse` ENABLE KEYS */;


--
-- Definition of table `tblprof`
--

DROP TABLE IF EXISTS `tblprof`;
CREATE TABLE `tblprof` (
  `id` text NOT NULL,
  `course` text NOT NULL,
  `yrsec` text NOT NULL,
  `fname` text NOT NULL,
  `gname` text NOT NULL,
  `mname` text NOT NULL,
  `gender` text NOT NULL,
  `cstatus` text NOT NULL,
  `bdate` text NOT NULL,
  `nationality` text NOT NULL,
  `tribe` text NOT NULL,
  `street` text NOT NULL,
  `brgy` text NOT NULL,
  `town` text NOT NULL,
  `prov` text NOT NULL,
  `zcode` text NOT NULL,
  `hphone` text NOT NULL,
  `email` text NOT NULL,
  `name` text NOT NULL,
  `relation` text NOT NULL,
  `contact` text NOT NULL
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

--
-- Dumping data for table `tblprof`
--

/*!40000 ALTER TABLE `tblprof` DISABLE KEYS */;
INSERT INTO `tblprof` (`id`,`course`,`yrsec`,`fname`,`gname`,`mname`,`gender`,`cstatus`,`bdate`,`nationality`,`tribe`,`street`,`brgy`,`town`,`prov`,`zcode`,`hphone`,`email`,`name`,`relation`,`contact`) VALUES 
 ('15306','BSCS','4A','BISQUERA','EZEKIEL','AGALOOS','MALE','SINGLE','01/30/78','FILIPINO','GADDANG','REGIDOR','SAN NICOLAS','BAYOMBONG','NUEVA VIZCAYA','3700','12345','mhelron_luv@yahoo.com','','','');
/*!40000 ALTER TABLE `tblprof` ENABLE KEYS */;


--
-- Definition of table `tblusers`
--

DROP TABLE IF EXISTS `tblusers`;
CREATE TABLE `tblusers` (
  `username` text NOT NULL,
  `password` text NOT NULL,
  `fname` text NOT NULL,
  `gname` text NOT NULL,
  `mname` text NOT NULL,
  `cpnum` text NOT NULL,
  `userlevel` text NOT NULL
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

--
-- Dumping data for table `tblusers`
--

/*!40000 ALTER TABLE `tblusers` DISABLE KEYS */;
INSERT INTO `tblusers` (`username`,`password`,`fname`,`gname`,`mname`,`cpnum`,`userlevel`) VALUES 
 ('mhel','mhel','bisquera','mhelo','danao','12345678911','admin');
/*!40000 ALTER TABLE `tblusers` ENABLE KEYS */;




/*!40101 SET SQL_MODE=@OLD_SQL_MODE */;
/*!40014 SET FOREIGN_KEY_CHECKS=@OLD_FOREIGN_KEY_CHECKS */;
/*!40014 SET UNIQUE_CHECKS=@OLD_UNIQUE_CHECKS */;
/*!40101 SET CHARACTER_SET_CLIENT=@OLD_CHARACTER_SET_CLIENT */;
/*!40101 SET CHARACTER_SET_RESULTS=@OLD_CHARACTER_SET_RESULTS */;
/*!40101 SET COLLATION_CONNECTION=@OLD_COLLATION_CONNECTION */;
/*!40101 SET CHARACTER_SET_CLIENT=@OLD_CHARACTER_SET_CLIENT */;
