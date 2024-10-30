/*
 Navicat Premium Data Transfer

 Source Server         : DHU考核系统
 Source Server Type    : SQLite
 Source Server Version : 3035005
 Source Schema         : main

 Target Server Type    : SQLite
 Target Server Version : 3035005
 File Encoding         : 65001

 Date: 30/10/2024 11:16:10
*/

PRAGMA foreign_keys = false;

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
  CONSTRAINT "dept_id" FOREIGN KEY ("dept_id") REFERENCES "department" ("dept_id") ON DELETE CASCADE ON UPDATE CASCADE,
  CONSTRAINT "index_id" FOREIGN KEY ("index_id") REFERENCES "index" ("index_id") ON DELETE CASCADE ON UPDATE CASCADE
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
-- Table structure for index
-- ----------------------------
DROP TABLE IF EXISTS "index";
CREATE TABLE "index" (
  "id" integer NOT NULL,
  "index_code" text NOT NULL,
  "index_name" TEXT,
  "index_type" TEXT,
  "weight1" real,
  "weight2" real,
  "enable_sensitivity" numeric,
  "sensitivity" real,
  PRIMARY KEY ("id")
);

-- ----------------------------
-- Table structure for index_duty
-- ----------------------------
DROP TABLE IF EXISTS "index_duty";
CREATE TABLE "index_duty" (
  "id" integer NOT NULL,
  "manager_id" integer NOT NULL,
  "index_id" integer NOT NULL,
  "enable_assessment" numeric,
  PRIMARY KEY ("id"),
  CONSTRAINT "manager_id" FOREIGN KEY ("manager_id") REFERENCES "manager" ("id") ON DELETE CASCADE ON UPDATE CASCADE,
  CONSTRAINT "index_id" FOREIGN KEY ("index_id") REFERENCES "index" ("id") ON DELETE CASCADE ON UPDATE CASCADE
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
