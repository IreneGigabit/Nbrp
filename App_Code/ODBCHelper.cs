using System;
using System.Collections.Generic;
using System.Web;
using System.Data;
using System.Data.Odbc;

/// <summary>
/// 資料庫操作類別
/// </summary>
public class ODBCHelper : IDisposable
{
    private OdbcConnection _conn = null;
	private OdbcTransaction _tran = null;
	private OdbcCommand _cmd = null;
	public string ConnString { get; set; }
	private bool _debug = false;
	private bool _isTran = true;

	public ODBCHelper(string connectionString) : this(connectionString, true) { }

    public ODBCHelper(string connectionString, bool isTransaction) {
		//this._debug = showDebugStr;
		this.ConnString = connectionString;
		this._isTran = isTransaction;

		this._conn = new OdbcConnection(this.ConnString);
		_conn.Open();

		if (this._isTran) {
			this._tran = _conn.BeginTransaction();
			this._cmd = new OdbcCommand("", _conn, _tran);
		} else {
			this._cmd = new OdbcCommand("", _conn);
		}
	}

    #region +ODBCHelper Debug(bool showDebugStr)
    public ODBCHelper Debug(bool showDebugStr) {
		this._debug = showDebugStr;
		return this;
	} 
	#endregion

	#region +void Dispose()
	public void Dispose() {
		this._conn.Close(); this._conn.Dispose();
		this._cmd.Dispose();
		if (this._tran != null) this._tran.Dispose();

		GC.SuppressFinalize(this);
	} 
	#endregion

	#region +void Commit()
	public void Commit() {
		if (this._tran != null) _tran.Commit();
	} 
	#endregion

	#region +void RollBack()
	public void RollBack() {
		if (this._tran != null) _tran.Rollback();
	} 
	#endregion

    #region 執行查詢，取得OdbcDataReader +OdbcDataReader ExecuteReader(string commandText)
    /// <summary>
    /// 執行查詢，取得OdbcDataReader；OdbcDataReader使用後須Close，否則會Lock(強烈建議使用using)。
	/// </summary>
	public OdbcDataReader ExecuteReader(string commandText) {
		if (this._debug) {
			HttpContext.Current.Response.Write(commandText + "<HR>");
		}
		this._cmd.CommandText = commandText;
        OdbcDataReader dr = this._cmd.ExecuteReader();

		return dr;
	} 
	#endregion

	#region 執行T-SQL，並傳回受影響的資料筆數 +int ExecuteNonQuery(string commandText)
	/// <summary>
	/// 執行T-SQL，並傳回受影響的資料筆數。
	/// </summary>
	public int ExecuteNonQuery(string commandText) {
		if (this._debug) {
			HttpContext.Current.Response.Write(commandText + "<HR>");
		}
		this._cmd.CommandText = commandText;
		return this._cmd.ExecuteNonQuery();
	} 
	#endregion

	#region 執行查詢，取得第一行第一欄資料，會忽略其他的資料行或資料列 +object ExecuteScalar(string commandText)
	/// <summary>
	/// 執行查詢，取得第一行第一欄資料，會忽略其他的資料行或資料列。
	/// </summary>
	public object ExecuteScalar(string commandText) {
		if (this._debug) {
			HttpContext.Current.Response.Write(commandText + "<HR>");
		}
		this._cmd.CommandText = commandText;
		return this._cmd.ExecuteScalar();
	} 
	#endregion

	#region 執行查詢，並傳回DataTable +void DataTable(string commandText, DataTable dt)
	/// <summary>
	/// 執行查詢，並傳回DataTable。
	/// </summary>
	public void DataTable(string commandText, DataTable dt) {
		if (this._debug) {
			HttpContext.Current.Response.Write(commandText + "<HR>");
		}
        using (OdbcDataAdapter adapter = new OdbcDataAdapter(commandText, this._conn)) {
			if (this._isTran) {
				adapter.SelectCommand.Transaction = this._tran;
			}
			adapter.Fill(dt);
		}
	} 
	#endregion

	#region 執行查詢，並傳回DataSet +void DataSet(string commandText, DataSet ds)
	/// <summary>
	/// 執行查詢，並傳回DataSet。
	/// </summary>
	public void DataSet(string commandText, DataSet ds) {
		if (this._debug) {
			HttpContext.Current.Response.Write(commandText + "<HR>");
		}
        using (OdbcDataAdapter adapter = new OdbcDataAdapter(commandText, this._conn)) {
			if (this._isTran) {
				adapter.SelectCommand.Transaction = this._tran;
			}
			//DataSet ds = new DataSet();
			adapter.Fill(ds);
		}
	}
	#endregion
}