CREATE DATABASE  IF NOT EXISTS `ao-new` /*!40100 DEFAULT CHARACTER SET utf8mb4 COLLATE utf8mb4_0900_ai_ci */ /*!80016 DEFAULT ENCRYPTION='N' */;
USE `ao-new`;
-- MySQL dump 10.13  Distrib 8.0.25, for Win64 (x86_64)
--
-- Host: localhost    Database: ao-new
-- ------------------------------------------------------
-- Server version	8.0.25

/*!40101 SET @OLD_CHARACTER_SET_CLIENT=@@CHARACTER_SET_CLIENT */;
/*!40101 SET @OLD_CHARACTER_SET_RESULTS=@@CHARACTER_SET_RESULTS */;
/*!40101 SET @OLD_COLLATION_CONNECTION=@@COLLATION_CONNECTION */;
/*!50503 SET NAMES utf8 */;
/*!40103 SET @OLD_TIME_ZONE=@@TIME_ZONE */;
/*!40103 SET TIME_ZONE='+00:00' */;
/*!40014 SET @OLD_UNIQUE_CHECKS=@@UNIQUE_CHECKS, UNIQUE_CHECKS=0 */;
/*!40014 SET @OLD_FOREIGN_KEY_CHECKS=@@FOREIGN_KEY_CHECKS, FOREIGN_KEY_CHECKS=0 */;
/*!40101 SET @OLD_SQL_MODE=@@SQL_MODE, SQL_MODE='NO_AUTO_VALUE_ON_ZERO' */;
/*!40111 SET @OLD_SQL_NOTES=@@SQL_NOTES, SQL_NOTES=0 */;

--
-- Table structure for table `account`
--

DROP TABLE IF EXISTS `account`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!50503 SET character_set_client = utf8mb4 */;
CREATE TABLE `account` (
  `id` int unsigned NOT NULL AUTO_INCREMENT,
  `email` varchar(100) NOT NULL,
  `password` varchar(32) NOT NULL,
  `first_name` varchar(30) NOT NULL,
  `last_name` varchar(30) NOT NULL,
  `validated` tinyint NOT NULL,
  `recovery_code` varchar(8) DEFAULT NULL,
  `validation_code` varchar(8) DEFAULT NULL,
  `last_ip` varchar(16) NOT NULL,
  `previous_ip` varchar(16) NOT NULL,
  `is_banned` tinyint NOT NULL,
  `ban_gm_id` int unsigned DEFAULT NULL,
  `ban_reason` varchar(200) DEFAULT NULL,
  `last_login_at` timestamp NULL DEFAULT NULL,
  `bank_gold` int NOT NULL,
  PRIMARY KEY (`id`),
  UNIQUE KEY `email_UNIQUE` (`email`),
  KEY `fk_account_character1_idx` (`ban_gm_id`)
) ENGINE=InnoDB AUTO_INCREMENT=18 DEFAULT CHARSET=utf8mb3;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Table structure for table `account_vault`
--

DROP TABLE IF EXISTS `account_vault`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!50503 SET character_set_client = utf8mb4 */;
CREATE TABLE `account_vault` (
  `account_id` int unsigned NOT NULL,
  `object_id` int unsigned NOT NULL,
  `position` tinyint unsigned NOT NULL,
  `amount` int unsigned NOT NULL,
  PRIMARY KEY (`account_id`,`position`),
  KEY `fk_character_inventory_object1_idx` (`object_id`),
  KEY `fk_account_vault_account1_idx` (`account_id`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb3;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Table structure for table `achievement`
--

DROP TABLE IF EXISTS `achievement`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!50503 SET character_set_client = utf8mb4 */;
CREATE TABLE `achievement` (
  `id` int unsigned NOT NULL AUTO_INCREMENT,
  `achievement_name` varchar(80) NOT NULL,
  `description` varchar(250) NOT NULL,
  PRIMARY KEY (`id`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb3;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Table structure for table `character`
--

DROP TABLE IF EXISTS `character`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!50503 SET character_set_client = utf8mb4 */;
CREATE TABLE `character` (
  `id` int unsigned NOT NULL AUTO_INCREMENT,
  `account_id` int unsigned NOT NULL,
  `character_name` varchar(40) NOT NULL,
  `deleted_at` timestamp NULL DEFAULT NULL,
  `created_at` timestamp NOT NULL DEFAULT CURRENT_TIMESTAMP,
  `last_ip` varchar(16) NOT NULL,
  `previous_ip` varchar(16) NOT NULL,
  `last_login_at` timestamp NULL DEFAULT CURRENT_TIMESTAMP,
  `last_logout_at` timestamp NULL DEFAULT NULL,
  `level` int NOT NULL,
  `exp` int NOT NULL,
  `genre_id` int NOT NULL,
  `race_id` int NOT NULL,
  `class_id` int NOT NULL,
  `city_id` int unsigned NOT NULL,
  `description` varchar(80) NOT NULL,
  `gold` int NOT NULL,
  `free_skillpoints` int NOT NULL,
  `pos_map` int NOT NULL,
  `pos_x` int NOT NULL,
  `pos_y` int NOT NULL,
  `body_id` int NOT NULL,
  `head_id` int NOT NULL,
  `wear_id` int NOT NULL DEFAULT '0',
  `weapon_id` int NOT NULL DEFAULT '0',
  `helmet_id` int NOT NULL DEFAULT '0',
  `shield_id` int NOT NULL DEFAULT '0',
  `ship_object_id` int NOT NULL,
  `heading` int NOT NULL,
  `min_hp` int NOT NULL,
  `max_hp` int NOT NULL,
  `min_man` int NOT NULL,
  `max_man` int NOT NULL,
  `min_sta` int NOT NULL,
  `max_sta` int NOT NULL,
  `min_hunger` int NOT NULL,
  `max_hunger` int NOT NULL,
  `min_thirst` int NOT NULL,
  `max_thirst` int NOT NULL,
  `min_hit` int NOT NULL,
  `max_hit` int NOT NULL,
  `invent_level` int NOT NULL,
  `vault_level` int NOT NULL,
  `guild_id` int unsigned DEFAULT NULL,
  `return_map` int NOT NULL,
  `return_x` int NOT NULL,
  `return_y` int NOT NULL,
  `delete_code` varchar(8) DEFAULT NULL,
  `faction` int NOT NULL,
  `faction_reward` int NOT NULL,
  `faction_entry_at` timestamp NULL DEFAULT NULL,
  `faction_entry_users` int NOT NULL,
  `killed_npcs` int NOT NULL,
  `killed_users` int NOT NULL,
  `killed_citizens` int NOT NULL,
  `killed_criminals` int NOT NULL,
  `deaths_npcs` int NOT NULL,
  `deaths_users` int NOT NULL,
  `deaths` int NOT NULL,
  `steps` bigint NOT NULL,
  `is_poisoned` tinyint NOT NULL,
  `is_incinerated` tinyint NOT NULL,
  `is_sailing` tinyint NOT NULL,
  `is_paralyzed` tinyint NOT NULL,
  `is_silenced` tinyint NOT NULL,
  `is_mounted` tinyint NOT NULL,
  `is_logged` tinyint NOT NULL,
  `chat_global` tinyint NOT NULL,
  `chat_combate` tinyint NOT NULL,
  `fishing_points` int NOT NULL,
  `elo` int NOT NULL,
  `time_played` bigint NOT NULL,
  `is_published` tinyint NOT NULL,
  `is_locked` tinyint NOT NULL,
  `jail_time` int NOT NULL,
  `silence_time` int NOT NULL,
  `banned_penalty_id` int unsigned DEFAULT NULL,
  PRIMARY KEY (`id`),
  KEY `fk_character_account_idx` (`account_id`),
  KEY `fk_character_city1_idx` (`city_id`),
  KEY `fk_character_guild1_idx` (`guild_id`),
  KEY `fk_character_character_penalty1_idx` (`banned_penalty_id`)
) ENGINE=InnoDB AUTO_INCREMENT=35 DEFAULT CHARSET=utf8mb3;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Table structure for table `character_achievement`
--

DROP TABLE IF EXISTS `character_achievement`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!50503 SET character_set_client = utf8mb4 */;
CREATE TABLE `character_achievement` (
  `character_id` int unsigned NOT NULL,
  `achievement_id` int unsigned NOT NULL,
  `created_at` timestamp NOT NULL DEFAULT CURRENT_TIMESTAMP,
  PRIMARY KEY (`character_id`,`achievement_id`),
  KEY `fk_character_achievement_character1_idx` (`character_id`),
  KEY `fk_character_achievement_achievement1_idx` (`achievement_id`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb3;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Table structure for table `character_event_log`
--

DROP TABLE IF EXISTS `character_event_log`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!50503 SET character_set_client = utf8mb4 */;
CREATE TABLE `character_event_log` (
  `id` bigint unsigned NOT NULL AUTO_INCREMENT,
  `character_id` int unsigned NOT NULL,
  `event_type_id` int unsigned NOT NULL,
  `created_at` timestamp NOT NULL DEFAULT CURRENT_TIMESTAMP,
  `description` varchar(250) NOT NULL,
  `entity_id` int unsigned DEFAULT NULL,
  `amount` int unsigned DEFAULT NULL,
  `value` int unsigned DEFAULT NULL,
  PRIMARY KEY (`id`),
  KEY `fk_character_log_event_character1_idx` (`character_id`)
) ENGINE=InnoDB AUTO_INCREMENT=42 DEFAULT CHARSET=utf8mb3;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Table structure for table `character_inventory`
--

DROP TABLE IF EXISTS `character_inventory`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!50503 SET character_set_client = utf8mb4 */;
CREATE TABLE `character_inventory` (
  `character_id` int unsigned NOT NULL,
  `object_id` int unsigned NOT NULL,
  `position` tinyint unsigned NOT NULL,
  `amount` int unsigned NOT NULL,
  `equipped` tinyint unsigned NOT NULL,
  PRIMARY KEY (`position`,`character_id`),
  KEY `fk_character_inventory_character1_idx` (`character_id`),
  KEY `fk_character_inventory_object1_idx` (`object_id`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb3;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Table structure for table `character_penalty`
--

DROP TABLE IF EXISTS `character_penalty`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!50503 SET character_set_client = utf8mb4 */;
CREATE TABLE `character_penalty` (
  `id` int unsigned NOT NULL AUTO_INCREMENT,
  `character_id` int unsigned NOT NULL,
  `created_at` timestamp NOT NULL DEFAULT CURRENT_TIMESTAMP,
  `gm_id` int unsigned DEFAULT NULL,
  `description` varchar(250) NOT NULL,
  `jail_time` int unsigned NOT NULL,
  `silence_time` int NOT NULL DEFAULT '0',
  `type` tinyint NOT NULL,
  `unban_at` timestamp NULL DEFAULT NULL,
  PRIMARY KEY (`id`),
  KEY `fk_character_penaly_character1_idx` (`character_id`) /*!80000 INVISIBLE */,
  KEY `fk_character_penalty_gm_idx` (`gm_id`)
) ENGINE=InnoDB AUTO_INCREMENT=17 DEFAULT CHARSET=utf8mb3;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Table structure for table `character_quest`
--

DROP TABLE IF EXISTS `character_quest`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!50503 SET character_set_client = utf8mb4 */;
CREATE TABLE `character_quest` (
  `character_id` int unsigned NOT NULL,
  `quest_id` int unsigned NOT NULL,
  `completed` tinyint NOT NULL,
  `data` text,
  `created_at` timestamp NOT NULL DEFAULT CURRENT_TIMESTAMP,
  `completed_at` timestamp NULL DEFAULT NULL,
  PRIMARY KEY (`character_id`,`quest_id`),
  KEY `fk_character_quest_character1_idx` (`character_id`),
  KEY `fk_character_quest_quest1_idx` (`quest_id`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb3;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Table structure for table `character_skill`
--

DROP TABLE IF EXISTS `character_skill`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!50503 SET character_set_client = utf8mb4 */;
CREATE TABLE `character_skill` (
  `character_id` int unsigned NOT NULL,
  `skill_id` int unsigned NOT NULL,
  `value` tinyint NOT NULL,
  `assigned` tinyint NOT NULL,
  PRIMARY KEY (`character_id`,`skill_id`),
  KEY `fk_character_skill_character1_idx` (`character_id`),
  KEY `fk_character_skill_skill1_idx` (`skill_id`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb3;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Table structure for table `character_spell`
--

DROP TABLE IF EXISTS `character_spell`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!50503 SET character_set_client = utf8mb4 */;
CREATE TABLE `character_spell` (
  `character_id` int unsigned NOT NULL,
  `spell_id` int unsigned NOT NULL,
  `position` tinyint NOT NULL,
  `created_at` timestamp NOT NULL DEFAULT CURRENT_TIMESTAMP,
  PRIMARY KEY (`character_id`,`spell_id`),
  KEY `fk_character_spell_character1_idx` (`character_id`),
  KEY `fk_character_spell_spell1_idx` (`spell_id`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb3;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Table structure for table `character_vault`
--

DROP TABLE IF EXISTS `character_vault`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!50503 SET character_set_client = utf8mb4 */;
CREATE TABLE `character_vault` (
  `character_id` int unsigned NOT NULL,
  `object_id` int unsigned NOT NULL,
  `position` tinyint unsigned NOT NULL,
  `amount` int unsigned NOT NULL,
  PRIMARY KEY (`position`,`character_id`),
  KEY `fk_character_inventory_character1_idx` (`character_id`),
  KEY `fk_character_inventory_object1_idx` (`object_id`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb3;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Table structure for table `city`
--

DROP TABLE IF EXISTS `city`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!50503 SET character_set_client = utf8mb4 */;
CREATE TABLE `city` (
  `id` int unsigned NOT NULL AUTO_INCREMENT,
  `city_name` varchar(50) NOT NULL,
  PRIMARY KEY (`id`)
) ENGINE=InnoDB AUTO_INCREMENT=2 DEFAULT CHARSET=utf8mb3;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Table structure for table `guild`
--

DROP TABLE IF EXISTS `guild`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!50503 SET character_set_client = utf8mb4 */;
CREATE TABLE `guild` (
  `id` int unsigned NOT NULL AUTO_INCREMENT,
  `guild_name` varchar(30) NOT NULL,
  `created_at` timestamp NOT NULL DEFAULT CURRENT_TIMESTAMP,
  PRIMARY KEY (`id`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb3;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Table structure for table `guild_member`
--

DROP TABLE IF EXISTS `guild_member`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!50503 SET character_set_client = utf8mb4 */;
CREATE TABLE `guild_member` (
  `guild_id` int unsigned NOT NULL,
  `character_id` int unsigned NOT NULL,
  `created_at` timestamp NOT NULL,
  `vote` int DEFAULT NULL,
  PRIMARY KEY (`guild_id`,`character_id`),
  KEY `fk_guild_has_character_guild1_idx` (`guild_id`),
  KEY `fk_guild_has_character_character1_idx` (`character_id`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb3;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Table structure for table `guild_relationship`
--

DROP TABLE IF EXISTS `guild_relationship`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!50503 SET character_set_client = utf8mb4 */;
CREATE TABLE `guild_relationship` (
  `guild1_id` int unsigned NOT NULL,
  `guild2_id` int unsigned NOT NULL,
  `accepted` tinyint NOT NULL,
  `status` tinyint NOT NULL,
  PRIMARY KEY (`guild1_id`,`guild2_id`),
  KEY `fk_guild_relationship_guild1_idx` (`guild1_id`),
  KEY `fk_guild_relationship_guild2_idx` (`guild2_id`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb3;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Table structure for table `guild_request`
--

DROP TABLE IF EXISTS `guild_request`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!50503 SET character_set_client = utf8mb4 */;
CREATE TABLE `guild_request` (
  `guild_id` int unsigned NOT NULL,
  `character_id` int unsigned NOT NULL,
  `created_at` timestamp NOT NULL DEFAULT CURRENT_TIMESTAMP,
  PRIMARY KEY (`guild_id`,`character_id`),
  KEY `fk_guild_request_guild1_idx` (`guild_id`),
  KEY `fk_guild_request_character1_idx` (`character_id`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb3;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Table structure for table `house`
--

DROP TABLE IF EXISTS `house`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!50503 SET character_set_client = utf8mb4 */;
CREATE TABLE `house` (
  `id` int unsigned NOT NULL AUTO_INCREMENT,
  `account_id` int unsigned DEFAULT NULL,
  `house_name` varchar(50) NOT NULL,
  `city_id` int unsigned NOT NULL,
  `assgined_at` timestamp NULL DEFAULT NULL,
  `door` int unsigned NOT NULL,
  `object_index` int NOT NULL,
  PRIMARY KEY (`id`),
  KEY `fk_house_account1_idx` (`account_id`),
  KEY `fk_house_city1_idx` (`city_id`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb3;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Table structure for table `login_log`
--

DROP TABLE IF EXISTS `login_log`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!50503 SET character_set_client = utf8mb4 */;
CREATE TABLE `login_log` (
  `id` int unsigned NOT NULL AUTO_INCREMENT,
  `account_id` int unsigned NOT NULL,
  `character_id` int unsigned DEFAULT NULL,
  `ip` varchar(16) NOT NULL,
  `created_at` timestamp NOT NULL DEFAULT CURRENT_TIMESTAMP,
  `ended_at` timestamp NULL DEFAULT NULL,
  PRIMARY KEY (`id`),
  KEY `fk_login_log_account1_idx` (`account_id`),
  KEY `fk_login_log_character1_idx` (`character_id`)
) ENGINE=InnoDB AUTO_INCREMENT=883 DEFAULT CHARSET=utf8mb3;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Table structure for table `market_bid`
--

DROP TABLE IF EXISTS `market_bid`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!50503 SET character_set_client = utf8mb4 */;
CREATE TABLE `market_bid` (
  `id` int unsigned NOT NULL AUTO_INCREMENT,
  `market_character_id` int unsigned NOT NULL,
  `value` double NOT NULL,
  `created_at` timestamp NOT NULL DEFAULT CURRENT_TIMESTAMP,
  `account_id` int unsigned NOT NULL,
  `comission_payed_at` timestamp NULL DEFAULT NULL,
  `payment_id` int unsigned DEFAULT NULL,
  `character1_id` int unsigned DEFAULT NULL,
  `character2_id` int unsigned DEFAULT NULL,
  `character3_id` int unsigned DEFAULT NULL,
  `accepted` tinyint NOT NULL,
  PRIMARY KEY (`id`),
  KEY `fk_market_bid_market_character1_idx` (`market_character_id`),
  KEY `fk_market_bid_account1_idx` (`account_id`),
  KEY `fk_market_bid_payment1_idx` (`payment_id`),
  KEY `fk_market_bid_character1_idx` (`character1_id`),
  KEY `fk_market_bid_character2_idx` (`character2_id`),
  KEY `fk_market_bid_character3_idx` (`character3_id`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb3;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Table structure for table `market_character`
--

DROP TABLE IF EXISTS `market_character`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!50503 SET character_set_client = utf8mb4 */;
CREATE TABLE `market_character` (
  `id` int unsigned NOT NULL AUTO_INCREMENT,
  `character_id` int unsigned NOT NULL,
  `seller_account_id` int unsigned NOT NULL,
  `created_at` timestamp NOT NULL DEFAULT CURRENT_TIMESTAMP,
  `initial_value` double NOT NULL,
  `current_value` double NOT NULL,
  `bid_ends_at` timestamp NULL DEFAULT NULL,
  `deleted_at` timestamp NULL DEFAULT NULL,
  `buyer_account_id` int unsigned NOT NULL,
  `fixed_price` tinyint NOT NULL,
  `allow_exchange` tinyint NOT NULL,
  `receiver_name` varchar(150) NOT NULL,
  `receiver_account` varchar(250) NOT NULL,
  `is_private` tinyint NOT NULL,
  `description` text NOT NULL,
  PRIMARY KEY (`id`),
  KEY `fk_market_character_character1_idx` (`character_id`),
  KEY `fk_market_character_account1_idx` (`seller_account_id`),
  KEY `fk_market_character_account2_idx` (`buyer_account_id`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb3;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Table structure for table `object`
--

DROP TABLE IF EXISTS `object`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!50503 SET character_set_client = utf8mb4 */;
CREATE TABLE `object` (
  `id` int unsigned NOT NULL AUTO_INCREMENT,
  `object_name` varchar(50) NOT NULL,
  PRIMARY KEY (`id`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb3;
/*!40101 SET character_set_client = @saved_cs_client */;
--
-- Table structure for table `payment`
--

DROP TABLE IF EXISTS `payment`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!50503 SET character_set_client = utf8mb4 */;
CREATE TABLE `payment` (
  `id` int unsigned NOT NULL AUTO_INCREMENT,
  `account_id` int unsigned NOT NULL,
  `created_at` timestamp NOT NULL DEFAULT CURRENT_TIMESTAMP,
  `transaction_id` varchar(45) NOT NULL,
  PRIMARY KEY (`id`),
  KEY `fk_payment_account1_idx` (`account_id`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb3;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Table structure for table `quest`
--

DROP TABLE IF EXISTS `quest`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!50503 SET character_set_client = utf8mb4 */;
CREATE TABLE `quest` (
  `id` int unsigned NOT NULL AUTO_INCREMENT,
  `quest_name` varchar(50) NOT NULL,
  PRIMARY KEY (`id`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb3;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Table structure for table `skill`
--

DROP TABLE IF EXISTS `skill`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!50503 SET character_set_client = utf8mb4 */;
CREATE TABLE `skill` (
  `id` int unsigned NOT NULL AUTO_INCREMENT,
  `skill_name` varchar(50) NOT NULL,
  PRIMARY KEY (`id`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb3;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Table structure for table `spell`
--

DROP TABLE IF EXISTS `spell`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!50503 SET character_set_client = utf8mb4 */;
CREATE TABLE `spell` (
  `id` int unsigned NOT NULL AUTO_INCREMENT,
  `spell_name` varchar(50) NOT NULL,
  `mana` int NOT NULL,
  `skills` int NOT NULL,
  PRIMARY KEY (`id`)
) ENGINE=InnoDB AUTO_INCREMENT=5 DEFAULT CHARSET=utf8mb3;
/*!40101 SET character_set_client = @saved_cs_client */;