package com.dha;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.Base64;
import java.util.Iterator;
import java.util.Properties;

import org.apache.log4j.Logger;
import org.apache.log4j.PropertyConfigurator;
import javax.security.auth.Subject;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.filenet.api.admin.ChoiceList;
import com.filenet.api.admin.ClassDefinition;
import com.filenet.api.admin.CodeModule;
import com.filenet.api.admin.PropertyTemplate;
import com.filenet.api.admin.TableDefinition;
import com.filenet.api.collection.AccessPermissionList;
import com.filenet.api.collection.IndependentObjectSet;
import com.filenet.api.constants.AccessType;
import com.filenet.api.constants.RefreshMode;
import com.filenet.api.core.Connection;
import com.filenet.api.core.CustomObject;
import com.filenet.api.core.Document;
import com.filenet.api.core.Domain;
import com.filenet.api.core.Factory;
import com.filenet.api.core.Folder;
import com.filenet.api.core.ObjectStore;
import com.filenet.api.core.WorkflowDefinition;
import com.filenet.api.events.Event;
import com.filenet.api.events.EventAction;
import com.filenet.api.events.Subscription;
import com.filenet.api.query.SearchSQL;
import com.filenet.api.query.SearchScope;
import com.filenet.api.security.AccessPermission;
import com.filenet.api.util.Id;
import com.filenet.api.util.UserContext;

public class MultiAdmin {
	private static final Logger logger = Logger.getLogger(MultiAdmin.class);
	private static String[] adminIds;
	private static String[] adminPasswords;
	private static String[] Objecttorenames;

	static {

		try {
			Properties props = new Properties();
			InputStream inputstream = MultiAdmin.class.getResourceAsStream("/com/dha/log4j.properties");
			props.load(inputstream);
			PropertyConfigurator.configure(props);
			logger.info(props);
			adminIds = props.getProperty("admin.ids").split(", ");
			adminPasswords = props.getProperty("admin.passwords").split(", ");
			Objecttorenames = props.getProperty("os").split(", ");
		} catch (Exception e) {
			logger.info(e);
		}

	}
	private static Connection conn = null;
	private static UserContext uc = null;
	private static String URI = "https://ecmcpetst.dohms.gov.ae/wsi/FNCEWS40MTOM";
//  private static	String URI = "https://ecmcpeprd.dohms.gov.ae/wsi/FNCEWS40MTOM";

	public static String logspath = "C:\\Users\\" + System.getProperty("user.name") + "\\" + "DHAMultiAdmin";

	public static void getConn(String userName, String password) {

		try {
			if (uc != null) {
				uc.popSubject();
				uc = null;
				conn = null;
			}
			if (conn == null) {

				conn = Factory.Connection.getConnection(URI);
				Subject subject = UserContext.createSubject(conn, userName, password, null);
				uc = UserContext.get();
				uc.pushSubject(subject);
				logger.info("the connection to----->   " + conn.getURI() + "    is established by ---> " + userName);
			}

		} catch (Exception e) {

			logger.info(e);
		}
	}

	public static void getthelist(String osname) {
//	getConn();

		String uname = "FNP8Admin";
		String pass = "QXByaWwyMDIwQDEyMw==";
		Base64.Decoder decoder = Base64.getUrlDecoder();
		// Decoding URl
		String password = new String(decoder.decode(pass));

		getConn(uname, password);
		logger.info(1);
		Domain domain = Factory.Domain.fetchInstance(conn, null, null);
		logger.info(1);
		ObjectStore os = Factory.ObjectStore.fetchInstance(domain, osname, null);

		logger.info(1);

		logger.info("the process of generating the list for " + osname + " is started");
		XSSFWorkbook osbook = null;
		FileInputStream osfis = null;
		try {
			osfis = new FileInputStream(logspath + "\\" + osname + "--list" + ".xlsx");
			logger.info(1);
			osbook = new XSSFWorkbook(osfis);
		} catch (Exception e) {
			logger.info(e);
			osbook = new XSSFWorkbook();
			logger.info(2);
		}

		Row row = null;
		Sheet sheet = null;
		int idx = 0;
		try {
			sheet = osbook.getSheet("Document");
			sheet.getSheetName();
		} catch (Exception e1) {
			sheet = osbook.createSheet("Document");
		}

		// -------------------------------------- Document //
		// --------------------------------------//

		SearchSQL sql = new SearchSQL("SELECT * FROM  [Document]");
		SearchScope scope = new SearchScope(os);
		IndependentObjectSet rowSet1 = scope.fetchObjects(sql, null, null, true);
		Iterator<?> it = rowSet1.iterator();

		while (it.hasNext()) {
			Document prop = (Document) it.next();

			AccessPermissionList apl1 = prop.get_Permissions();
			Iterator<?> itr = apl1.iterator();
			boolean bool = false;
			while (itr.hasNext()) {
				AccessPermission ap = (AccessPermission) itr.next();

				if (ap.get_GranteeName().equals("ECM_OS_Admins_stg@dohms.gov.ae")) {
					bool = true;
					break;

				}

			}
			if (!bool) {
				logger.info("yes no security for Ecm_OS_Admins_stg group for ---" + prop.get_Name());
				String docid = prop.get_Id().toString();
				int rowOffset = idx;
				while (rowOffset > 1048575) {
					rowOffset = rowOffset - 1048576;
				}

				int cellOffset = idx / 1048576;
				row = sheet.getRow(rowOffset);
				if (row == null) {
					row = sheet.createRow(rowOffset);
				}
				Cell cell = row.createCell(cellOffset);
				cell.setCellValue(docid);
				logger.info(idx);
				idx++;
			}
			FileOutputStream fos;
			try {

				fos = new FileOutputStream(logspath + "\\" + osname + "--list" + ".xlsx");
				osbook.write(fos);
				fos.close();
			} catch (IOException e1) {
				e1.printStackTrace();
			}

		}

		// ----------------------------------- Documents are done //
		// ---------------------------------------//

		// -------------------------------------- Event //
		// --------------------------------------//

		row = null;
		sheet = null;
		idx = 0;
		try {
			sheet = osbook.getSheet("Event");
			sheet.getSheetName();
		} catch (Exception e1) {
			sheet = osbook.createSheet("Event");
		}

		sql = new SearchSQL("SELECT * FROM  [Event]");
		scope = new SearchScope(os);
		rowSet1 = scope.fetchObjects(sql, null, null, true);
		it = rowSet1.iterator();

		while (it.hasNext()) {
			Event prop = (Event) it.next();

			AccessPermissionList apl1 = prop.get_Permissions();
			Iterator<?> itr = apl1.iterator();
			boolean bool = false;
			while (itr.hasNext()) {
				AccessPermission ap = (AccessPermission) itr.next();

				if (ap.get_GranteeName().equals("ECM_OS_Admins_stg@dohms.gov.ae")) {
					bool = true;
					break;

				}

			}
			if (!bool) {
				logger.info("yes no security for Ecm_OS_Admins_stg group for ---" + prop.get_Name());
				String docid = prop.get_Id().toString();
				int rowOffset = idx;
				while (rowOffset > 1048575) {
					rowOffset = rowOffset - 1048576;
				}

				int cellOffset = idx / 1048576;
				row = sheet.getRow(rowOffset);
				if (row == null) {
					row = sheet.createRow(rowOffset);
				}
				Cell cell = row.createCell(cellOffset);
				cell.setCellValue(docid);
				logger.info(idx);
				idx++;
			}
			FileOutputStream fos;
			try {

				fos = new FileOutputStream(logspath + "\\" + osname + "--list" + ".xlsx");
				osbook.write(fos);
				fos.close();
			} catch (IOException e1) {
				e1.printStackTrace();
			}

		}

		// ----------------------------------- Events are done//
		// ---------------------------------------//

		// -------------------------------------- Folder //
		// --------------------------------------//

		row = null;
		sheet = null;
		idx = 0;
		try {
			sheet = osbook.getSheet("Folder");
			sheet.getSheetName();
		} catch (Exception e1) {
			sheet = osbook.createSheet("Folder");
		}

		sql = new SearchSQL("SELECT * FROM  [Folder]");
		scope = new SearchScope(os);
		rowSet1 = scope.fetchObjects(sql, null, null, true);
		it = rowSet1.iterator();

		while (it.hasNext()) {
			Folder prop = (Folder) it.next();

			AccessPermissionList apl1 = prop.get_Permissions();
			Iterator<?> itr = apl1.iterator();
			boolean bool = false;
			while (itr.hasNext()) {
				AccessPermission ap = (AccessPermission) itr.next();

				if (ap.get_GranteeName().equals("ECM_OS_Admins_stg@dohms.gov.ae")) {
					bool = true;
					break;

				}

			}
			if (!bool) {
				logger.info("yes no security for Ecm_OS_Admins_stg group for ---" + prop.get_Name());
				String docid = prop.get_Id().toString();
				int rowOffset = idx;
				while (rowOffset > 1048575) {
					rowOffset = rowOffset - 1048576;
				}

				int cellOffset = idx / 1048576;
				row = sheet.getRow(rowOffset);
				if (row == null) {
					row = sheet.createRow(rowOffset);
				}
				Cell cell = row.createCell(cellOffset);
				cell.setCellValue(docid);
				logger.info(idx);
				idx++;
			}
			FileOutputStream fos;
			try {

				fos = new FileOutputStream(logspath + "\\" + osname + "--list" + ".xlsx");
				osbook.write(fos);
				fos.close();
			} catch (IOException e1) {
				e1.printStackTrace();
			}

		}

		// ----------------------------------- Folders are done //
		// ---------------------------------------//

		// -------------------------------------- CustomObject //
		// --------------------------------------//

		row = null;
		sheet = null;
		idx = 0;
		try {
			sheet = osbook.getSheet("CustomObject");
			sheet.getSheetName();
		} catch (Exception e1) {
			sheet = osbook.createSheet("CustomObject");
		}

		sql = new SearchSQL("SELECT * FROM  [CustomObject]");
		scope = new SearchScope(os);
		rowSet1 = scope.fetchObjects(sql, null, null, true);
		it = rowSet1.iterator();

		while (it.hasNext()) {
			CustomObject prop = (CustomObject) it.next();

			AccessPermissionList apl1 = prop.get_Permissions();
			Iterator<?> itr = apl1.iterator();
			boolean bool = false;
			while (itr.hasNext()) {
				AccessPermission ap = (AccessPermission) itr.next();

				if (ap.get_GranteeName().equals("ECM_OS_Admins_stg@dohms.gov.ae")) {
					bool = true;
					break;

				}

			}
			if (!bool) {
				logger.info("yes no security for Ecm_OS_Admins_stg group for ---" + prop.get_Name());
				String docid = prop.get_Id().toString();
				int rowOffset = idx;
				while (rowOffset > 1048575) {
					rowOffset = rowOffset - 1048576;
				}

				int cellOffset = idx / 1048576;
				row = sheet.getRow(rowOffset);
				if (row == null) {
					row = sheet.createRow(rowOffset);
				}
				Cell cell = row.createCell(cellOffset);
				cell.setCellValue(docid);
				logger.info(idx);
				idx++;
			}
			FileOutputStream fos;
			try {

				fos = new FileOutputStream(logspath + "\\" + osname + "--list" + ".xlsx");
				osbook.write(fos);
				fos.close();
			} catch (IOException e1) {
				e1.printStackTrace();
			}

		}

		// ----------------------------------- CustomObjects are done //
		// ---------------------------------------//

		// -------------------------------------- ClassDefinition //
		// --------------------------------------//

		row = null;
		sheet = null;
		idx = 0;
		try {
			sheet = osbook.getSheet("ClassDefinition");
			sheet.getSheetName();
		} catch (Exception e1) {
			sheet = osbook.createSheet("ClassDefinition");
		}

		sql = new SearchSQL("SELECT * FROM  [ClassDefinition]");
		scope = new SearchScope(os);
		rowSet1 = scope.fetchObjects(sql, null, null, true);
		it = rowSet1.iterator();

		while (it.hasNext()) {
			ClassDefinition prop = (ClassDefinition) it.next();

			AccessPermissionList apl1 = prop.get_Permissions();
			Iterator<?> itr = apl1.iterator();
			boolean bool = false;
			while (itr.hasNext()) {
				AccessPermission ap = (AccessPermission) itr.next();

				if (ap.get_GranteeName().equals("ECM_OS_Admins_stg@dohms.gov.ae")) {
					bool = true;
					break;

				}

			}
			if (!bool) {
				logger.info("yes no security for Ecm_OS_Admins_stg group for ---" + prop.get_Name());
				String docid = prop.get_Id().toString();
				int rowOffset = idx;
				while (rowOffset > 1048575) {
					rowOffset = rowOffset - 1048576;
				}

				int cellOffset = idx / 1048576;
				row = sheet.getRow(rowOffset);
				if (row == null) {
					row = sheet.createRow(rowOffset);
				}
				Cell cell = row.createCell(cellOffset);
				cell.setCellValue(docid);
				logger.info(idx);
				idx++;
			}
			FileOutputStream fos;
			try {

				fos = new FileOutputStream(logspath + "\\" + osname + "--list" + ".xlsx");
				osbook.write(fos);
				fos.close();
			} catch (IOException e1) {
				e1.printStackTrace();
			}

		}

		// ----------------------------------- ClassDefinitions are done //
		// ---------------------------------------//

		// -------------------------------------- PropertyTemplate //
		// --------------------------------------//

		row = null;
		sheet = null;
		idx = 0;
		try {
			sheet = osbook.getSheet("PropertyTemplate");
			sheet.getSheetName();
		} catch (Exception e1) {
			sheet = osbook.createSheet("PropertyTemplate");
		}

		sql = new SearchSQL("SELECT * FROM  [PropertyTemplate]");
		scope = new SearchScope(os);
		rowSet1 = scope.fetchObjects(sql, null, null, true);
		it = rowSet1.iterator();

		while (it.hasNext()) {
			PropertyTemplate prop = (PropertyTemplate) it.next();

			AccessPermissionList apl1 = prop.get_Permissions();
			Iterator<?> itr = apl1.iterator();
			boolean bool = false;
			while (itr.hasNext()) {
				AccessPermission ap = (AccessPermission) itr.next();

				if (ap.get_GranteeName().equals("ECM_OS_Admins_stg@dohms.gov.ae")) {
					bool = true;
					break;

				}

			}
			if (!bool) {
				logger.info("yes no security for Ecm_OS_Admins_stg group for ---" + prop.get_Name());
				String docid = prop.get_Id().toString();
				int rowOffset = idx;
				while (rowOffset > 1048575) {
					rowOffset = rowOffset - 1048576;
				}

				int cellOffset = idx / 1048576;
				row = sheet.getRow(rowOffset);
				if (row == null) {
					row = sheet.createRow(rowOffset);
				}
				Cell cell = row.createCell(cellOffset);
				cell.setCellValue(docid);
				logger.info(idx);
				idx++;
			}
			FileOutputStream fos;
			try {

				fos = new FileOutputStream(logspath + "\\" + osname + "--list" + ".xlsx");
				osbook.write(fos);
				fos.close();
			} catch (IOException e1) {
				e1.printStackTrace();
			}

		}

		// ----------------------------------- PropertyTemplates are done //
		// ---------------------------------------//

		// -------------------------------------- EventAction //
		// --------------------------------------//

		row = null;
		sheet = null;
		idx = 0;
		try {
			sheet = osbook.getSheet("EventAction");
			sheet.getSheetName();
		} catch (Exception e1) {
			sheet = osbook.createSheet("EventAction");
		}

		sql = new SearchSQL("SELECT * FROM  [EventAction]");
		scope = new SearchScope(os);
		rowSet1 = scope.fetchObjects(sql, null, null, true);
		it = rowSet1.iterator();

		while (it.hasNext()) {
			EventAction prop = (EventAction) it.next();

			AccessPermissionList apl1 = prop.get_Permissions();
			Iterator<?> itr = apl1.iterator();
			boolean bool = false;
			while (itr.hasNext()) {
				AccessPermission ap = (AccessPermission) itr.next();

				if (ap.get_GranteeName().equals("ECM_OS_Admins_stg@dohms.gov.ae")) {
					bool = true;
					break;

				}

			}
			if (!bool) {
				logger.info("yes no security for Ecm_OS_Admins_stg group for ---" + prop.get_Name());
				String docid = prop.get_Id().toString();
				int rowOffset = idx;
				while (rowOffset > 1048575) {
					rowOffset = rowOffset - 1048576;
				}

				int cellOffset = idx / 1048576;
				row = sheet.getRow(rowOffset);
				if (row == null) {
					row = sheet.createRow(rowOffset);
				}
				Cell cell = row.createCell(cellOffset);
				cell.setCellValue(docid);
				logger.info(idx);
				idx++;
			}
			FileOutputStream fos;
			try {

				fos = new FileOutputStream(logspath + "\\" + osname + "--list" + ".xlsx");
				osbook.write(fos);
				fos.close();
			} catch (IOException e1) {
				e1.printStackTrace();
			}

		}

		// ----------------------------------- EventActions are done //
		// ---------------------------------------//

		// -------------------------------------- ChoiceList //
		// --------------------------------------//

		row = null;
		sheet = null;
		idx = 0;
		try {
			sheet = osbook.getSheet("ChoiceList");
			sheet.getSheetName();
		} catch (Exception e1) {
			sheet = osbook.createSheet("ChoiceList");
		}

		sql = new SearchSQL("SELECT * FROM  [ChoiceList]");
		scope = new SearchScope(os);
		rowSet1 = scope.fetchObjects(sql, null, null, true);
		it = rowSet1.iterator();

		while (it.hasNext()) {
			ChoiceList prop = (ChoiceList) it.next();

			AccessPermissionList apl1 = prop.get_Permissions();
			Iterator<?> itr = apl1.iterator();
			boolean bool = false;
			while (itr.hasNext()) {
				AccessPermission ap = (AccessPermission) itr.next();

				if (ap.get_GranteeName().equals("ECM_OS_Admins_stg@dohms.gov.ae")) {
					bool = true;
					break;

				}

			}
			if (!bool) {
				logger.info("yes no security for Ecm_OS_Admins_stg group for ---" + prop.get_Name());
				String docid = prop.get_Id().toString();
				int rowOffset = idx;
				while (rowOffset > 1048575) {
					rowOffset = rowOffset - 1048576;
				}

				int cellOffset = idx / 1048576;
				row = sheet.getRow(rowOffset);
				if (row == null) {
					row = sheet.createRow(rowOffset);
				}
				Cell cell = row.createCell(cellOffset);
				cell.setCellValue(docid);
				logger.info(idx);
				idx++;
			}
			FileOutputStream fos;
			try {

				fos = new FileOutputStream(logspath + "\\" + osname + "--list" + ".xlsx");
				osbook.write(fos);
				fos.close();
			} catch (IOException e1) {
				e1.printStackTrace();
			}

		}

		// ----------------------------------- ChoiceLists are done //
		// ---------------------------------------//

		// -------------------------------------- WorkflowDefinition //
		// --------------------------------------//

		row = null;
		sheet = null;
		idx = 0;
		try {
			sheet = osbook.getSheet("WorkflowDefinition");
			sheet.getSheetName();
		} catch (Exception e1) {
			sheet = osbook.createSheet("WorkflowDefinition");
		}

		sql = new SearchSQL("SELECT * FROM  [WorkflowDefinition]");
		scope = new SearchScope(os);
		rowSet1 = scope.fetchObjects(sql, null, null, true);
		it = rowSet1.iterator();

		while (it.hasNext()) {
			WorkflowDefinition prop = (WorkflowDefinition) it.next();

			AccessPermissionList apl1 = prop.get_Permissions();
			Iterator<?> itr = apl1.iterator();
			boolean bool = false;
			while (itr.hasNext()) {
				AccessPermission ap = (AccessPermission) itr.next();

				if (ap.get_GranteeName().equals("ECM_OS_Admins_stg@dohms.gov.ae")) {
					bool = true;
					break;

				}

			}
			if (!bool) {
				logger.info("yes no security for Ecm_OS_Admins_stg group for ---" + prop.get_Name());
				String docid = prop.get_Id().toString();
				int rowOffset = idx;
				while (rowOffset > 1048575) {
					rowOffset = rowOffset - 1048576;
				}

				int cellOffset = idx / 1048576;
				row = sheet.getRow(rowOffset);
				if (row == null) {
					row = sheet.createRow(rowOffset);
				}
				Cell cell = row.createCell(cellOffset);
				cell.setCellValue(docid);
				logger.info(idx);
				idx++;
			}
			FileOutputStream fos;
			try {

				fos = new FileOutputStream(logspath + "\\" + osname + "--list" + ".xlsx");
				osbook.write(fos);
				fos.close();
			} catch (IOException e1) {
				e1.printStackTrace();
			}

		}

		// ----------------------------------- WorkflowDefinitions are done //
		// ---------------------------------------//

		// -------------------------------------- Subscription //
		// --------------------------------------//

		row = null;
		sheet = null;
		idx = 0;
		try {
			sheet = osbook.getSheet("Subscription");
			sheet.getSheetName();
		} catch (Exception e1) {
			sheet = osbook.createSheet("Subscription");
		}

		sql = new SearchSQL("SELECT * FROM  [Subscription]");
		scope = new SearchScope(os);
		rowSet1 = scope.fetchObjects(sql, null, null, true);
		it = rowSet1.iterator();

		while (it.hasNext()) {
			Subscription prop = (Subscription) it.next();

			AccessPermissionList apl1 = prop.get_Permissions();
			Iterator<?> itr = apl1.iterator();
			boolean bool = false;
			while (itr.hasNext()) {
				AccessPermission ap = (AccessPermission) itr.next();

				if (ap.get_GranteeName().equals("ECM_OS_Admins_stg@dohms.gov.ae")) {
					bool = true;
					break;

				}

			}
			if (!bool) {
				logger.info("yes no security for Ecm_OS_Admins_stg group for ---" + prop.get_Name());
				String docid = prop.get_Id().toString();
				int rowOffset = idx;
				while (rowOffset > 1048575) {
					rowOffset = rowOffset - 1048576;
				}

				int cellOffset = idx / 1048576;
				row = sheet.getRow(rowOffset);
				if (row == null) {
					row = sheet.createRow(rowOffset);
				}
				Cell cell = row.createCell(cellOffset);
				cell.setCellValue(docid);
				logger.info(idx);
				idx++;
			}
			FileOutputStream fos;
			try {

				fos = new FileOutputStream(logspath + "\\" + osname + "--list" + ".xlsx");
				osbook.write(fos);
				fos.close();
			} catch (IOException e1) {
				e1.printStackTrace();
			}

		}

		// ----------------------------------- Subscriptions are done //
		// ---------------------------------------//

		// -------------------------------------- TableDefinition //
		// --------------------------------------//

		row = null;
		sheet = null;
		idx = 0;
		try {
			sheet = osbook.getSheet("TableDefinition");
			sheet.getSheetName();
		} catch (Exception e1) {
			sheet = osbook.createSheet("TableDefinition");
		}

		sql = new SearchSQL("SELECT * FROM  [TableDefinition]");
		scope = new SearchScope(os);
		rowSet1 = scope.fetchObjects(sql, null, null, true);
		it = rowSet1.iterator();

		while (it.hasNext()) {
			TableDefinition prop = (TableDefinition) it.next();

			AccessPermissionList apl1 = prop.get_Permissions();
			Iterator<?> itr = apl1.iterator();
			boolean bool = false;
			while (itr.hasNext()) {
				AccessPermission ap = (AccessPermission) itr.next();

				if (ap.get_GranteeName().equals("ECM_OS_Admins_stg@dohms.gov.ae")) {
					bool = true;
					break;

				}

			}
			if (!bool) {
				logger.info("yes no security for Ecm_OS_Admins_stg group for ---" + prop.get_Name());
				String docid = prop.get_Id().toString();
				int rowOffset = idx;
				while (rowOffset > 1048575) {
					rowOffset = rowOffset - 1048576;
				}

				int cellOffset = idx / 1048576;
				row = sheet.getRow(rowOffset);
				if (row == null) {
					row = sheet.createRow(rowOffset);
				}
				Cell cell = row.createCell(cellOffset);
				cell.setCellValue(docid);
				logger.info(idx);
				idx++;
			}
			FileOutputStream fos;
			try {

				fos = new FileOutputStream(logspath + "\\" + osname + "--list" + ".xlsx");
				osbook.write(fos);
				fos.close();
			} catch (IOException e1) {
				e1.printStackTrace();
			}

		}

		// ----------------------------------- TableDefinitions are done //
		// ---------------------------------------//

		// -------------------------------------- CodeModule //
		// --------------------------------------//

		row = null;
		sheet = null;
		idx = 0;
		try {
			sheet = osbook.getSheet("CodeModule");
			sheet.getSheetName();
		} catch (Exception e1) {
			sheet = osbook.createSheet("CodeModule");
		}

		sql = new SearchSQL("SELECT * FROM  [CodeModule]");
		scope = new SearchScope(os);
		rowSet1 = scope.fetchObjects(sql, null, null, true);
		it = rowSet1.iterator();

		while (it.hasNext()) {
			CodeModule prop = (CodeModule) it.next();

			AccessPermissionList apl1 = prop.get_Permissions();
			Iterator<?> itr = apl1.iterator();
			boolean bool = false;
			while (itr.hasNext()) {
				AccessPermission ap = (AccessPermission) itr.next();

				if (ap.get_GranteeName().equals("ECM_OS_Admins_stg@dohms.gov.ae")) {
					bool = true;
					break;

				}

			}
			if (!bool) {
				logger.info("yes no security for Ecm_OS_Admins_stg group for ---" + prop.get_Name());
				String docid = prop.get_Id().toString();
				int rowOffset = idx;
				while (rowOffset > 1048575) {
					rowOffset = rowOffset - 1048576;
				}

				int cellOffset = idx / 1048576;
				row = sheet.getRow(rowOffset);
				if (row == null) {
					row = sheet.createRow(rowOffset);
				}
				Cell cell = row.createCell(cellOffset);
				cell.setCellValue(docid);
				logger.info(idx);
				idx++;
			}
			FileOutputStream fos;
			try {

				fos = new FileOutputStream(logspath + "\\" + osname + "--list" + ".xlsx");
				osbook.write(fos);
				fos.close();
			} catch (IOException e1) {
				e1.printStackTrace();
			}

		}
		// ----------------------------------- CodeModules are done //
		// ---------------------------------------//
		logger.info("List generation for the " + osname + "is completed");
		uc.popSubject();

	}

	@SuppressWarnings({ "resource", "unchecked", "removal" })
	public static void MultiAdminSecurity(String os) {
		try {
			logger.info("in");

			for (int i = 0; i < adminIds.length; i++) {
				String username = adminIds[i];
				String pass = adminPasswords[i];

				Base64.Decoder decoder = Base64.getUrlDecoder();
				String dStr = new String(decoder.decode(pass));
				String password = dStr;
				getConn(username, password);
				Domain dom = null;
				ObjectStore OS = null;
				try {
					dom = Factory.Domain.fetchInstance(conn, null, null);
					OS = Factory.ObjectStore.fetchInstance(dom, os, null);
				} catch (Exception e) {
					// TODO: handle exception
					logger.info(e);
					continue;
				}

				logger.info("the security adding process for " + os + " is started");
				File osfile = new File(logspath + "\\" + os + "--list" + ".xlsx");
				FileInputStream osfis = new FileInputStream(osfile);

				XSSFWorkbook workbook = new XSSFWorkbook(osfis);
				int numSheets = workbook.getNumberOfSheets();
				logger.info(numSheets);
				for (int j = 0; j < numSheets; j++) {

					Sheet sheet1 = workbook.getSheetAt(j);
					String shname = sheet1.getSheetName();
					logger.info(shname);
					Iterator<Row> rowIterator = sheet1.iterator();
					while (rowIterator.hasNext()) {
						Row row1 = rowIterator.next();
						Iterator<Cell> cellIterator = row1.iterator();
						while (cellIterator.hasNext()) {
							Cell cell = cellIterator.next();
							String id = cell.toString();
							logger.info(id);
							if (shname.equals("Document") && id != "") {
								logger.info("doc" + id);

								try {
									Document doc = Factory.Document.fetchInstance(OS, new Id(id), null);
									AccessPermissionList apl = doc.get_Permissions();
									AccessPermission permission = Factory.AccessPermission.createInstance();
									permission.set_GranteeName("ECM_OS_Admins_stg");
									permission.set_AccessType(AccessType.ALLOW);
									permission.set_InheritableDepth(new Integer(0));
									permission.set_AccessMask(999415);
									apl.add(permission);
									doc.set_Permissions(apl);
									doc.save(RefreshMode.REFRESH);
									cell.setCellValue("");

								} catch (Exception e) {
									logger.info(e);
								}

							} else if (shname.equals("Event") && id != "") {
								logger.info("Event " + id);

								try {
									Event doc = Factory.Event.fetchInstance(OS, new Id(id), null);
									AccessPermissionList apl = doc.get_Permissions();
									AccessPermission permission = Factory.AccessPermission.createInstance();
									permission.set_GranteeName("ECM_OS_Admins_stg");
									permission.set_AccessType(AccessType.ALLOW);
									permission.set_InheritableDepth(new Integer(0));
									permission.set_AccessMask(995587);
									apl.add(permission);
									doc.set_Permissions(apl);
									doc.save(RefreshMode.REFRESH);
									cell.setCellValue("");

								} catch (Exception e) {
									logger.info(e);
								}

							} else if (shname.equals("Folder") && id != "") {
								logger.info("Folder " + id);

								try {
									Folder doc = Factory.Folder.fetchInstance(OS, new Id(id), null);
									AccessPermissionList apl = doc.get_Permissions();
									AccessPermission permission = Factory.AccessPermission.createInstance();
									permission.set_GranteeName("ECM_OS_Admins_stg");
									permission.set_AccessType(AccessType.ALLOW);
									permission.set_InheritableDepth(new Integer(0));
									permission.set_AccessMask(999415);
									apl.add(permission);
									doc.set_Permissions(apl);
									doc.save(RefreshMode.REFRESH);
									cell.setCellValue("");

								} catch (Exception e) {
									logger.info(e);
								}

							} else if (shname.equals("CustomObject") && id != "") {
								logger.info("Co " + id);

								try {
									CustomObject doc = Factory.CustomObject.fetchInstance(OS, new Id(id), null);
									AccessPermissionList apl = doc.get_Permissions();
									AccessPermission permission = Factory.AccessPermission.createInstance();
									permission.set_GranteeName("ECM_OS_Admins_stg");
									permission.set_AccessType(AccessType.ALLOW);
									permission.set_InheritableDepth(new Integer(0));
									permission.set_AccessMask(995603);
									apl.add(permission);
									doc.set_Permissions(apl);
									doc.save(RefreshMode.REFRESH);
									cell.setCellValue("");

								} catch (Exception e) {
									logger.info(e);
								}

							} else if (shname.equals("ClassDefinition") && id != "") {
								logger.info("cd" + id);
								try {
									ClassDefinition doc = Factory.ClassDefinition.fetchInstance(OS, new Id(id), null);

									AccessPermissionList apl = doc.get_Permissions();
									AccessPermission permission = Factory.AccessPermission.createInstance();
									permission.set_GranteeName("ECM_OS_Admins_stg");
									permission.set_AccessType(AccessType.ALLOW);
									permission.set_InheritableDepth(new Integer(0));
									permission.set_AccessMask(983827);
									apl.add(permission);
									doc.set_Permissions(apl);
									doc.save(RefreshMode.REFRESH);

								} catch (Exception e) {
									logger.info(e);
								}

							} else if (shname.equals("PropertyTemplate") && id != "") {
								logger.info("pt " + id);
								try {
									PropertyTemplate doc = Factory.PropertyTemplate.fetchInstance(OS, new Id(id), null);

									AccessPermissionList apl = doc.get_Permissions();
									AccessPermission permission = Factory.AccessPermission.createInstance();
									permission.set_GranteeName("ECM_OS_Admins_stg");
									permission.set_AccessType(AccessType.ALLOW);
									permission.set_InheritableDepth(new Integer(0));
									permission.set_AccessMask(983827);
									apl.add(permission);
									doc.set_Permissions(apl);
									doc.save(RefreshMode.REFRESH);
									cell.setCellValue("");

								} catch (Exception e) {
									logger.info(e);
								}

							} else if (shname.equals("EventAction") && id != "") {
								logger.info("EA " + id);
								try {
									EventAction doc = Factory.EventAction.fetchInstance(OS, new Id(id), null);

									AccessPermissionList apl = doc.get_Permissions();
									AccessPermission permission = Factory.AccessPermission.createInstance();
									permission.set_GranteeName("ECM_OS_Admins_stg");
									permission.set_AccessType(AccessType.ALLOW);
									permission.set_InheritableDepth(new Integer(0));
									permission.set_AccessMask(983827);
									apl.add(permission);
									doc.set_Permissions(apl);
									doc.save(RefreshMode.REFRESH);
									cell.setCellValue("");

								} catch (Exception e) {
									logger.info(e);
								}

							} else if (shname.equals("ChoiceList") && id != "") {
								logger.info("cl " + id);
								try {
									ChoiceList doc = Factory.ChoiceList.fetchInstance(OS, new Id(id), null);

									AccessPermissionList apl = doc.get_Permissions();
									AccessPermission permission = Factory.AccessPermission.createInstance();
									permission.set_GranteeName("ECM_OS_Admins_stg");
									permission.set_AccessType(AccessType.ALLOW);
									permission.set_InheritableDepth(new Integer(0));
									permission.set_AccessMask(983827);
									apl.add(permission);
									doc.set_Permissions(apl);
									doc.save(RefreshMode.REFRESH);
									cell.setCellValue("");

								} catch (Exception e) {
									logger.info(e);
								}

							} else if (shname.equals("WorkflowDefinition") && id != "") {
								logger.info("wd " + id);
								try {
									WorkflowDefinition doc = Factory.WorkflowDefinition.fetchInstance(OS, new Id(id),
											null);

									AccessPermissionList apl = doc.get_Permissions();
									AccessPermission permission = Factory.AccessPermission.createInstance();
									permission.set_GranteeName("ECM_OS_Admins_stg");
									permission.set_AccessType(AccessType.ALLOW);
									permission.set_InheritableDepth(new Integer(0));
									permission.set_AccessMask(998903);
									apl.add(permission);
									doc.set_Permissions(apl);
									doc.save(RefreshMode.REFRESH);
									cell.setCellValue("");

								} catch (Exception e) {
									logger.info(e);
								}
							} else if (shname.equals("Subscription") && id != "") {
								logger.info("sub " + id);
								try {
									Subscription doc = Factory.Subscription.fetchInstance(OS, new Id(id), null);

									AccessPermissionList apl = doc.get_Permissions();
									AccessPermission permission = Factory.AccessPermission.createInstance();
									permission.set_GranteeName("ECM_OS_Admins_stg");
									permission.set_AccessType(AccessType.ALLOW);
									permission.set_InheritableDepth(new Integer(0));
									permission.set_AccessMask(998903);
									apl.add(permission);
									doc.set_Permissions(apl);
									doc.save(RefreshMode.REFRESH);
									cell.setCellValue("");

								} catch (Exception e) {
									logger.info(e);
								}

							} else if (shname.equals("TableDefinition") && id != "") {
								logger.info("tb " + id);

								try {
									TableDefinition doc = Factory.TableDefinition.fetchInstance(OS, new Id(id), null);

									AccessPermissionList apl = doc.get_Permissions();
									AccessPermission permission = Factory.AccessPermission.createInstance();
									permission.set_GranteeName("ECM_OS_Admins_stg");
									permission.set_AccessType(AccessType.ALLOW);
									permission.set_InheritableDepth(new Integer(0));
									permission.set_AccessMask(983827);
									apl.add(permission);
									doc.set_Permissions(apl);
									doc.save(RefreshMode.REFRESH);
									cell.setCellValue("");

								} catch (Exception e) {
									logger.info(e);
								}
							} else if (shname.equals("CodeModule") && id != "") {
								logger.info("cm");

								try {
									CodeModule doc = Factory.CodeModule.fetchInstance(OS, new Id(id), null);

									AccessPermissionList apl = doc.get_Permissions();
									AccessPermission permission = Factory.AccessPermission.createInstance();
									permission.set_GranteeName("ECM_OS_Admins_stg");
									permission.set_AccessType(AccessType.ALLOW);
									permission.set_InheritableDepth(new Integer(0));
									permission.set_AccessMask(998903);
									apl.add(permission);
									doc.set_Permissions(apl);
									doc.save(RefreshMode.REFRESH);
									cell.setCellValue("");

								} catch (Exception e) {
									logger.info(e);
								}

							}

						}
					}

				}
				FileOutputStream osfos = new FileOutputStream(osfile);
				workbook.write(osfos);
				osfos.close();
				logger.info("the adding security for the " + OS.get_DisplayName() + "is fininshed");
				osfis.close();
			}

		} catch (Exception e) {
			// TODO: handle exception
			logger.info(e);
		}

	}

	@SuppressWarnings("resource")
	public static void main(String[] args) {

		try {

			for (int i = 0; i < Objecttorenames.length; i++) {

				String osvalue = Objecttorenames[i];
				logger.info(osvalue);
				getthelist(osvalue);
				MultiAdminSecurity(osvalue);

			}

		} catch (Exception e) {
			// TODO: handle exception
			logger.info(e);
		}

	}
}
