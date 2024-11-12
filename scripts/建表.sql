/*
 Navicat Premium Data Transfer

 Source Server         : DHU考核系统
 Source Server Type    : SQLite
 Source Server Version : 3035005
 Source Schema         : main

 Target Server Type    : SQLite
 Target Server Version : 3035005
 File Encoding         : 65001

 Date: 12/11/2024 18:26:59
*/

PRAGMA FOREIGN_KEYS=ON;

-- ----------------------------
-- Table structure for _groups_old_20241112
-- ----------------------------
DROP TABLE IF EXISTS "_groups_old_20241112";
CREATE TABLE "_groups_old_20241112" (
  "id" INTEGER NOT NULL,
  "group_name" TEXT NOT NULL,
  "l_bound" TEXT,
  "r_bound" TEXT,
  PRIMARY KEY ("id")
);

-- ----------------------------
-- Table structure for _groups_old_20241112_1
-- ----------------------------
DROP TABLE IF EXISTS "_groups_old_20241112_1";
CREATE TABLE "_groups_old_20241112_1" (
  "id" INTEGER NOT NULL,
  "index_id" INTEGER NOT NULL,
  "group_name" TEXT NOT NULL,
  "l_bound" TEXT,
  "r_bound" TEXT,
  PRIMARY KEY ("id")
);

-- ----------------------------
-- Table structure for _indexes_old_20241107
-- ----------------------------
DROP TABLE IF EXISTS "_indexes_old_20241107";
CREATE TABLE "_indexes_old_20241107" (
  "id" integer NOT NULL,
  "identifier_id" INTEGER,
  "secondary_identifier" integer,
  "index_name" TEXT,
  "index_type" TEXT,
  "weight1" real,
  "weight2" real,
  "enable_sensitivity" integer,
  "sensitivity" real,
  PRIMARY KEY ("id"),
  CONSTRAINT "identifier_id" FOREIGN KEY ("identifier_id") REFERENCES "index_identifier" ("id") ON DELETE CASCADE ON UPDATE CASCADE
);

-- ----------------------------
-- Table structure for completion
-- ----------------------------
DROP TABLE IF EXISTS "completion";
CREATE TABLE "completion" (
  "id" integer NOT NULL,
  "dept_id" integer NOT NULL,
  "year" integer NOT NULL,
  "index_id" integer NOT NULL,
  "target" integer,
  "completed" integer,
  PRIMARY KEY ("id"),
  CONSTRAINT "dept_id" FOREIGN KEY ("dept_id") REFERENCES "department" ("id") ON DELETE CASCADE ON UPDATE CASCADE,
  CONSTRAINT "index_id" FOREIGN KEY ("index_id") REFERENCES "indexes" ("id") ON DELETE CASCADE ON UPDATE CASCADE
);

-- ----------------------------
-- Table structure for department
-- ----------------------------
DROP TABLE IF EXISTS "department";
CREATE TABLE "department" (
  "id" INTEGER NOT NULL,
  "dept_code" text NOT NULL,
  "dept_name" text,
  PRIMARY KEY ("id")
);

-- ----------------------------
-- Table structure for dept_annual_info
-- ----------------------------
DROP TABLE IF EXISTS "dept_annual_info";
CREATE TABLE "dept_annual_info" (
  "id" integer NOT NULL,
  "dept_id" integer NOT NULL,
  "year" integer NOT NULL,
  "dept_population" integer,
  "dept_punishment" real,
  "dept_group" TEXT,
  PRIMARY KEY ("id"),
  CONSTRAINT "dept_id" FOREIGN KEY ("dept_id") REFERENCES "department" ("id") ON DELETE CASCADE ON UPDATE CASCADE
);

-- ----------------------------
-- Table structure for group_completion
-- ----------------------------
DROP TABLE IF EXISTS "group_completion";
CREATE TABLE "group_completion" (
  "id" integer NOT NULL,
  "group_id" integer NOT NULL,
  "year" integer NOT NULL,
  "index_id" integer NOT NULL,
  "target" integer,
  "completed" integer,
  PRIMARY KEY ("id"),
  CONSTRAINT "group_id" FOREIGN KEY ("group_id") REFERENCES "groups" ("id") ON DELETE CASCADE ON UPDATE CASCADE,
  CONSTRAINT "index_id" FOREIGN KEY ("index_id") REFERENCES "indexes" ("id") ON DELETE CASCADE ON UPDATE CASCADE
);

-- ----------------------------
-- Table structure for groups
-- ----------------------------
DROP TABLE IF EXISTS "groups";
CREATE TABLE "groups" (
  "id" INTEGER NOT NULL,
  "index_id" INTEGER NOT NULL,
  "group_name" TEXT NOT NULL,
  "l_bound" TEXT,
  "r_bound" TEXT,
  PRIMARY KEY ("id"),
  CONSTRAINT "index_id" FOREIGN KEY ("index_id") REFERENCES "indexes" ("id") ON DELETE CASCADE ON UPDATE CASCADE
);

-- ----------------------------
-- Table structure for index_duty
-- ----------------------------
DROP TABLE IF EXISTS "index_duty";
CREATE TABLE "index_duty" (
  "id" integer NOT NULL,
  "manager_id" integer NOT NULL,
  "index_id" integer NOT NULL,
  "enable_assessment" integer,
  PRIMARY KEY ("id"),
  CONSTRAINT "manager_id" FOREIGN KEY ("manager_id") REFERENCES "manager" ("id") ON DELETE CASCADE ON UPDATE CASCADE,
  CONSTRAINT "index_id" FOREIGN KEY ("index_id") REFERENCES "indexes" ("id") ON DELETE CASCADE ON UPDATE CASCADE
);

-- ----------------------------
-- Table structure for index_identifier
-- ----------------------------
DROP TABLE IF EXISTS "index_identifier";
CREATE TABLE "index_identifier" (
  "id" INTEGER NOT NULL,
  "identifier_name" TEXT,
  PRIMARY KEY ("id")
);

-- ----------------------------
-- Table structure for indexes
-- ----------------------------
DROP TABLE IF EXISTS "indexes";
CREATE TABLE "indexes" (
  "id" integer NOT NULL,
  "identifier_id" INTEGER,
  "secondary_identifier" integer,
  "tertiary_identifier" TEXT,
  "index_name" TEXT,
  "index_type" TEXT,
  "weight1" real,
  "weight2" real,
  "enable_sensitivity" integer,
  "sensitivity" real,
  PRIMARY KEY ("id"),
  CONSTRAINT "identifier_id" FOREIGN KEY ("identifier_id") REFERENCES "index_identifier" ("id") ON DELETE CASCADE ON UPDATE CASCADE
);

-- ----------------------------
-- Table structure for manager
-- ----------------------------
DROP TABLE IF EXISTS "manager";
CREATE TABLE "manager" (
  "id" INTEGER NOT NULL,
  "manager_code" text NOT NULL,
  "manager_name" TEXT,
  PRIMARY KEY ("id")
);

PRAGMA foreign_keys = true;
