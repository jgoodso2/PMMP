using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Data.SqlClient;
using SvcLookupTable;
using System.ServiceModel;
using WCFHelpers;
using System.Diagnostics;

namespace Repository
{
    /// <summary>
    /// 
    /// </summary>
    public class LookupPSIDataSet : IPSIDataSet
    {
        private const string RBS_CON_STRING = "RBSConnectionString";
        private string connectionString;

        public LookupPSIDataSet()
        {
            try
            {
                connectionString = System.Configuration.ConfigurationManager.ConnectionStrings[RBS_CON_STRING].ConnectionString;
            }
            catch (Exception ex)
            {
                Utility.WriteLog("Connection string for ParallonRBSLoad not present in App Config. Please add one",EventLogEntryType.Error);
                throw ex;
            }
        }

        public DataSet GetDataSet()
        {
            try
            {
                //Read lookup tables from the project server
                DataSet ds = DataRepository.ReadLookupTables();
                Utility.WriteLog("Get Data Set From project server succeeded", EventLogEntryType.Information);
                //Test GitHub Check in changes
                return ds;
            }
            catch(Exception ex)
            {
                Utility.WriteLog("Error in Reading the lookup table from the Project server = " + ex.Message, EventLogEntryType.Error);
                throw ex;
            }
           
        }

        public void Update(DataSet delta)
        {
            try
            {
                DataRepository.UpdateLookupTables((LookupTableDataSet)delta);
                Utility.WriteLog("Update to project server completely succeeded", EventLogEntryType.Information);
            }
            catch (Exception ex)
            {
                Utility.WriteLog("Error in Update = " + ex.Message, EventLogEntryType.Error);
            }
        }
        public DataSet GetChanges()
        {
            try
            {
                //Read DataSet from the database updated by SSIS Package
                SqlConnection sqlCon = new SqlConnection(connectionString);
                SqlDataAdapter adapter = new SqlDataAdapter();
                DataSet ds = new DataSet();
                SqlCommand sqlCommand = new SqlCommand();
                sqlCommand.Connection = sqlCon;
                sqlCommand.CommandText = "select * FROM [RBS_Values_Latest] s ORDER BY s.SortIndex asc";
                adapter.SelectCommand = sqlCommand;
                adapter.Fill(ds);
                Utility.WriteLog("Get Changes from the SSIS Database succeeded", EventLogEntryType.Information);
                return ds;
            }
            catch (Exception ex)
            {
                Utility.WriteLog("Error in Gettting changes from the SSIS Database = " + ex.Message, EventLogEntryType.Error);
            }
            return new DataSet();
        }


        public DataSet GetDelta(DataSet source, DataSet changes)
        {
            try
            {
                LookupTableDataSet lookUpSource = (LookupTableDataSet)source;
                LookupTableDataSet lookUpSourceCopy = (LookupTableDataSet)source.Copy();
                LookupTableDataSet delta = new LookupTableDataSet();
                // Build tow lists of Look up DTO's with one list repsenting the root node(all node with rowlevel =1)
                //and the second list containing all the child nodes( all node at row levels other  thna 1)
                List<List<LookupDTO>> lookups = BuildLookUpObject(changes);
                //For each root element
                foreach (LookupDTO lookup in lookups[0])
                {
                    // check to ensure it is a root element so as to check does not have any parent node
                    if (lookUpSourceCopy.LookupTableTrees.Any(t => t.Field<Guid?>("LT_PARENT_STRUCT_UID") == (Guid?)lookup.ID) == false)
                    {
                        //If the root node not already existing  add a new root node
                        if (lookUpSourceCopy.LookupTableTrees.Any(t => t.Field<Guid?>("LT_STRUCT_UID") == (Guid?)lookup.ID) == false)
                        {
                            AddTreeNode(lookup, lookUpSourceCopy);
                        }
                        //else Modify the root node
                        else
                        {
                            Modify(lookUpSourceCopy.LookupTableTrees.First(t => t.Field<Guid?>("LT_STRUCT_UID") == (Guid?)lookup.ID), lookup);
                        }
                    }
                }
                //For each child node
                foreach (LookupDTO lookup in lookups[1])
                {
                    //If child node not already existing add a new node
                    if (lookUpSourceCopy.LookupTableTrees.Any(t => t.Field<Guid?>("LT_STRUCT_UID") == (Guid?)lookup.ID) == false)
                    {
                        AddChildNode(lookup, lookUpSourceCopy);
                    }
                    //else modify the existing node
                    else
                    {
                        LookupTableDataSet.LookupTableTreesRow rowNode = lookUpSourceCopy.LookupTableTrees.First(t => t.RowState != DataRowState.Deleted && t.Field<Guid?>("LT_STRUCT_UID") == (Guid?)lookup.ID);
                        Modify(rowNode, lookup);
                    }
                }
                // Traverse through all child node of each root node and if it is unchanged delete it
                foreach (LookupDTO lookup in lookups[0])
                {
                    //Delete all child rows that are unchanged
                    DeleteAllChilds(lookup.ID, lookUpSourceCopy);
                }

                if (lookUpSourceCopy.HasChanges(DataRowState.Added))
                    delta.Merge(
                        (SvcLookupTable.LookupTableDataSet)lookUpSourceCopy.GetChanges(DataRowState.Added), true);

                if (lookUpSourceCopy.HasChanges(DataRowState.Modified))
                    delta.Merge(
                        (SvcLookupTable.LookupTableDataSet)lookUpSourceCopy.GetChanges(DataRowState.Modified), true);

                if (lookUpSourceCopy.HasChanges(DataRowState.Deleted))
                    delta.Merge(
                        (SvcLookupTable.LookupTableDataSet)lookUpSourceCopy.GetChanges(DataRowState.Deleted), true);
                Utility.WriteLog("Get Delta from the Source and Destination Database succeeded", EventLogEntryType.Information);
                return delta;
            }
            catch (Exception ex)
            {
                Utility.WriteLog("Error in Get Delta from the Source and Destination Database =" + ex.Message, EventLogEntryType.Error);
            }
            return new DataSet();
        }

        private void Modify(LookupTableDataSet.LookupTableTreesRow row, LookupDTO lookup)
        {
            try
            {
                if (row.RowState == DataRowState.Deleted)
                {
                    row.RejectChanges();
                }
                row.LT_PARENT_STRUCT_UID = lookup.ParentID;
                row.LT_VALUE_SORT_INDEX = lookup.SortIndex;
                row.LT_VALUE_TEXT = lookup.Text;
                //row.LT_VALUE_FULL = lookup.DotNotation;
                row.LT_VALUE_DESC = string.Empty;
                row.LT_UID = new Guid(Constants.LOOKUP_ENTITY_ID);
            }
            catch (Exception ex)
            {
                Utility.WriteLog("Error in Modifying a Lookup table tree row =" + ex.Message, EventLogEntryType.Error);
            }
        }

        /// <summary>
        /// Add Child node to the lookup Tree
        /// </summary>
        /// <param name="lookup"></param>
        /// <param name="lookUpSourceCopy"></param>
        private void AddChildNode(LookupDTO lookup, LookupTableDataSet lookUpSourceCopy)
        {
            try
            {
                //If its parent node is present for the node
                if (lookUpSourceCopy.LookupTableTrees.Any(t => t.Field<Guid?>("LT_STRUCT_UID") == (Guid?)lookup.ParentID) == true)
                {
                    // Add or modify depending on whether the node itself already exists
                    if (lookUpSourceCopy.LookupTableTrees.Any(t => t.Field<Guid?>("LT_STRUCT_UID") == (Guid?)lookup.ID) == false)
                    {
                        AddTreeNode(lookup, lookUpSourceCopy);
                    }
                    else
                    {
                        LookupTableDataSet.LookupTableTreesRow row = lookUpSourceCopy.LookupTableTrees.First(t => t.Field<Guid?>("LT_STRUCT_UID") == (Guid?)lookup.ID);
                        Modify(row, lookup);
                    }
                }
                    //Else add up the parent node 
                else
                {
                    // If aprent node does not already exist in the lookup Tree
                    if (lookUpSourceCopy.LookupTableTrees.Any(t => t.Field<Guid?>("LT_STRUCT_UID") == (Guid?)lookup.ParentID) == false)
                    {
                        // Add Parent Node
                        if (lookup.ParentNode != null)
                        {
                            AddTreeNode(lookup.ParentNode, lookUpSourceCopy);
                        }
                    }
                    else
                    {
                        //Else modify the Parent Node
                        if (lookup.ParentNode != null)
                        {
                            LookupTableDataSet.LookupTableTreesRow row = lookUpSourceCopy.LookupTableTrees.First(t => t.Field<Guid?>("LT_STRUCT_UID") == (Guid?)lookup.ParentID);
                            Modify(row, lookup.ParentNode);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Utility.WriteLog("Error in Adding a Lookup table tree row =" + ex.Message, EventLogEntryType.Error);
            }
        }

        /// <summary>
        /// Add Tree Node to the look up Tree
        /// </summary>
        /// <param name="lookup"></param>
        /// <param name="lookUpSourceCopy"></param>
        private void AddTreeNode(LookupDTO lookup, LookupTableDataSet lookUpSourceCopy)
        {
            try
            {
                LookupTableDataSet.LookupTableTreesRow row = lookUpSourceCopy.LookupTableTrees.NewLookupTableTreesRow();
                row.LT_STRUCT_UID = lookup.ID;
                if (lookup.ParentID != Guid.Empty)
                {
                    row.LT_PARENT_STRUCT_UID = lookup.ParentID;
                }

                row.LT_VALUE_SORT_INDEX = lookup.SortIndex;
                row.LT_VALUE_TEXT = lookup.Text;
                row.LT_VALUE_FULL = lookup.DotNotation;
                row.LT_VALUE_DESC = string.Empty;
                row.LT_UID = new Guid(Constants.LOOKUP_ENTITY_ID);
                lookUpSourceCopy.LookupTableTrees.AddLookupTableTreesRow(row);
            }
            catch (Exception ex)
            {
                Utility.WriteLog("Error in Adding a Lookup table tree row =" + ex.Message, EventLogEntryType.Error);
            }
        }

        /// <summary>
        /// Delete all child nodes for a given root element
        /// </summary>
        /// <param name="guid"></param>
        /// <param name="lookUpSourceCopy"></param>
        private void DeleteAllChilds(Guid guid, LookupTableDataSet lookUpSourceCopy)
        {
            try
            {
                if (lookUpSourceCopy.LookupTableTrees.Any(t => t.RowState != DataRowState.Deleted && t.Field<Guid?>("LT_PARENT_STRUCT_UID") == (Guid?)guid) == false)
                    return;
                EnumerableRowCollection<LookupTableDataSet.LookupTableTreesRow> rows = lookUpSourceCopy.LookupTableTrees.Where(t => t.RowState != DataRowState.Deleted && t.Field<Guid?>("LT_PARENT_STRUCT_UID") == (Guid?)guid);

                foreach (LookupTableDataSet.LookupTableTreesRow row in rows)
                {
                    DeleteAllChilds(row.LT_STRUCT_UID, lookUpSourceCopy);
                    if (row.RowState == DataRowState.Unchanged || row.RowState == DataRowState.Detached)
                    {
                        row.Delete();
                    }
                }
            }
            catch (Exception ex)
            {
                Utility.WriteLog("Error in Deleting a Lookup table tree row =" + ex.Message, EventLogEntryType.Error);
            }
        }


        /// <summary>
        /// Build a lookup Data List of Root Node and Child node lists into two lists returned
        /// </summary>
        /// <param name="ds"></param>
        /// <returns></returns>
        private List<List<LookupDTO>> BuildLookUpObject(DataSet ds)
        {
            try
            {
                //List of Root Node and Child node lists into two lists returned
                List<List<LookupDTO>> nodes = new List<List<LookupDTO>>();
                //List of root Nodes
                List<LookupDTO> rootListDto = new List<LookupDTO>();
                //List of child Nodes
                List<LookupDTO> childistDto = new List<LookupDTO>();
                //Use a Dictionary so that you can add a parent node to a child node when you need it
                Dictionary<Guid, LookupDTO> dictLookup = new Dictionary<Guid, LookupDTO>();
                if (ds.Tables.Count > 0)
                {
                    //For each row read from the SSIS package
                    foreach (DataRow row in ds.Tables[0].Rows)
                    {
                        //Create a lookup DTO
                        LookupDTO lookupDTO = MapLookupRow(row);
                        //If it is a root node add to the root node list
                        if (lookupDTO.ParentID == Guid.Empty)
                        {
                            rootListDto.Add(lookupDTO);
                            dictLookup.Add(lookupDTO.ID, lookupDTO);
                        }
                        //else add to the child node list
                        else
                        {
                            if (dictLookup.ContainsKey(lookupDTO.ParentID))
                            {
                                lookupDTO.ParentNode = dictLookup[lookupDTO.ParentID];
                            }
                            childistDto.Add(lookupDTO);
                            dictLookup.Add(lookupDTO.ID, lookupDTO);
                        }
                    }
                }
                return new List<List<LookupDTO>> { rootListDto, childistDto };
            }
            catch (Exception ex)
            {
                Utility.WriteLog("Error in Building DTO for  a Lookup table tree row =" + ex.Message, EventLogEntryType.Error);
            }
            return new List<List<LookupDTO>>();
        }
        /// <summary>
        /// Maps a single row from a SSIS data to a DTO object
        /// </summary>
        /// <param name="row"></param>
        /// <returns></returns>
        private LookupDTO MapLookupRow(DataRow row)
        {
            try
            {
                LookupDTO dto = new LookupDTO();
                if (row != null)
                {

                    dto.ID = row["GeneratedGUID"] != System.DBNull.Value ? new Guid(row["GeneratedGUID"].ToString()) : Guid.Empty;
                    dto.ParentID = row["ParentGeneratedGUID"] != System.DBNull.Value ? new Guid(row["ParentGeneratedGUID"].ToString()) : Guid.Empty;
                    dto.LastLoad = row["LastLoad"] != System.DBNull.Value ? Convert.ToDateTime(row["LastLoad"].ToString()) : default(DateTime);
                    dto.ProcessingDate = row["ProcessingDate"] != System.DBNull.Value ? Convert.ToDateTime(row["ProcessingDate"].ToString()) : default(DateTime);
                    dto.RowLevel = row["RowLevel"] != System.DBNull.Value ? Convert.ToInt32(row["RowLevel"].ToString()) : 0;
                    dto.SortIndex = row["SortIndex"] != System.DBNull.Value ? Convert.ToInt32(row["SortIndex"].ToString()) : 0;
                    dto.Text = row["Item"] != System.DBNull.Value ? row["Item"].ToString() : string.Empty;
                    dto.DotNotation = row["DotNotation"] != System.DBNull.Value ? row["DotNotation"].ToString() : string.Empty;
                    dto.COID = row["COID"] != System.DBNull.Value ? row["COID"].ToString() : string.Empty;
                }
                return dto;
            }
            catch (Exception ex)
            {
                Utility.WriteLog("Error in Mapping DTO for  a Lookup table tree row =" + ex.Message, EventLogEntryType.Error);
            }
         return new LookupDTO();   
        }
    }
}
