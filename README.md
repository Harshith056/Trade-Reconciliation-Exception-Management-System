# Trade Reconciliation & Exception Management System

## Overview

This project simulates a capital markets operations workflow by automating trade reconciliation between trade and settlement records. The system identifies operational exceptions, generates reconciliation reports, and provides interactive Power BI dashboards for monitoring settlement performance and operational risk.

## Features

* Automated trade reconciliation
* Quantity mismatch detection
* Price mismatch detection
* Settlement failure identification
* Exception reporting
* KPI generation
* Power BI dashboard integration
* Financial market analytics

## Technology Stack

* Python
* Pandas
* Excel
* Power BI
* OpenPyXL

## Workflow

Trade Book → Settlement Book → Reconciliation Engine → Exception Reports → Power BI Dashboard

## Reports Generated

### Reconciliation Report

Contains all trade records and reconciliation outcomes.

### Exception Report

Contains only trades requiring investigation.

### KPI Report

Provides operational metrics including:

* Total Trades
* Matched Trades
* Break Rate
* Settlement Success Rate
* Quantity Mismatches
* Price Mismatches
* Settlement Failures

## Dashboard Pages

### Executive Summary

* Total Trades
* Matched Trades
* Break Rate
* Settlement Success Rate
* Exception Breakdown

### Exception Analysis

* Exception Distribution
* Exception Counts
* Investigation Queue

### Market Analytics

* Closing Price Trend
* Trading Volume Trend
* RSI Trend
* Volatility Trend

## Project Outcome

The system processes 50,000+ trade records and demonstrates reconciliation, reporting, exception management, and financial data analytics workflows commonly used in capital markets operations.
