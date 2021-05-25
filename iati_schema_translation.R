library(XML)

setwd("~/git/how-publishers-publish")

xpaths = c()

act_schema = "iati-activities-schema.xsd"
common_schema = "iati-common.xsd"

ELEMENT_SEP = "/"
ATTRIBUTE_SEP = "/@"
TEXT_SEL = "text()"

recursive_xpath_constructor = function(xpaths, ns, current_xpath, xml_root, element_name, element_type, complex_types){
  message(element_name)
  if(!is.null(element_type)){
    if(element_type %in% names(complex_types)){
      element = complex_types[[element_type]]
      possible_children = unlist(getNodeSet(
        doc = element,
        path = "./xsd:sequence/xsd:element",
        ns
      ))
      possible_attributes = unlist(getNodeSet(
        doc = element,
        path = "./xsd:attribute/@name",
        ns
      ))
    }else{
      element_selector = paste0("//xsd:element[@name='", element_name, "']")
      element = getNodeSet(xml_root, element_selector)[[1]]
      possible_children = unlist(getNodeSet(
        doc = element,
        path = "./xsd:complexType/xsd:sequence/xsd:element",
        ns
      ))
      possible_attributes = unlist(getNodeSet(
        doc = element,
        path = "./xsd:complexType/xsd:attribute/@name",
        ns
      ))
    }
  }else{
    element_selector = paste0("//xsd:element[@name='", element_name, "']")
    element = getNodeSet(xml_root, element_selector)[[1]]
    possible_children = unlist(getNodeSet(
      doc = element,
      path = "./xsd:complexType/xsd:sequence/xsd:element",
      ns
    ))
    possible_attributes = unlist(getNodeSet(
      doc = element,
      path = "./xsd:complexType/xsd:attribute/@name",
      ns
    ))
  }

  current_xpath = paste(current_xpath, element_name, sep=ELEMENT_SEP)

  for(possible_attribute in possible_attributes){
    attrib_xpath = paste(current_xpath, possible_attribute, sep=ATTRIBUTE_SEP)
    xpaths = c(xpaths, attrib_xpath)
  }
  if(length(possible_children) == 0 ){
    text_xpath = paste(current_xpath, TEXT_SEL, sep=ELEMENT_SEP)
    xpaths = c(xpaths, text_xpath)
  }
  for(possible_child in possible_children){
    element_name = unlist(getNodeSet(possible_child, "./@ref | ./@name"))
    element_type = unlist(getNodeSet(possible_child, "./@type"))
    element_simple = unlist(getNodeSet(possible_child, "./xsd:complexType/xsd:simpleContent", ns))
    if(length(element_simple) > 0 & length(possible_children)>0){
      text_xpath = paste(current_xpath, element_name, TEXT_SEL, sep=ELEMENT_SEP)
      xpaths = c(xpaths, text_xpath)
    }else{
      xpaths = recursive_xpath_constructor(xpaths, ns, current_xpath, xml_root, element_name, element_type, complex_types)
    }
  }
  return(xpaths)
}

current_xpath = ""
xml_doc = xmlParse(act_schema)
common_doc = xmlParse(common_schema)
xml_root = xmlRoot(xml_doc)
xml_root = addChildren(xml_root, getNodeSet(xmlRoot(common_doc),"xsd:element"))
xml_root = addChildren(xml_root, getNodeSet(xmlRoot(common_doc),"xsd:complexType"))
complex_types = getNodeSet(xmlRoot(common_doc),"xsd:complexType")
names(complex_types) = getNodeSet(xmlRoot(common_doc),"xsd:complexType/@name")

nsDefs <- xmlNamespaceDefinitions(xml_doc)
ns <- structure(sapply(nsDefs, function(x) x$uri), names = names(nsDefs))

root_element = getNodeSet(xml_root, "xsd:element[@name='iati-activities']")[[1]]
root_element_name = 'iati-activities'
current_xpath = paste(current_xpath, root_element_name, sep=ELEMENT_SEP)
possible_children = unlist(getNodeSet(
  doc = root_element, 
  path = "./xsd:complexType/xsd:sequence/xsd:element",
  ns
))
possible_attributes = unlist(getNodeSet(
  doc = root_element,
  path = "./xsd:complexType/xsd:attribute/@name",
  ns
))
for(possible_attribute in possible_attributes){
  attrib_xpath = paste(current_xpath, possible_attribute, sep=ATTRIBUTE_SEP)
  xpaths = c(xpaths, attrib_xpath)
}
for(possible_child in possible_children){
  element_name = unlist(getNodeSet(possible_child, "./@ref | ./@name"))
  element_type = unlist(getNodeSet(possible_child, "./@type"))
  xpaths = recursive_xpath_constructor(xpaths, ns, current_xpath, xml_root, element_name, element_type, complex_types)
}

writeLines(xpaths,"iati_schema_xpaths.txt")