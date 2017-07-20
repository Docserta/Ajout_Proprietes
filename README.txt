Macro de documentation d'attributs sur les parts et products et d'export

Modules:
-------
A_AjoutProprietes
	Permet d'importer une liste d'attributs déclarés dans un fichier texte
	sur tous les parts et product d'un assemblage

B_ExtractNom
	Permet d'exporter vers un fichier Excel les attributs des parts et products d'un assemblage
	en les regroupant par ensemble/sous ensemble et Details

C_ModifProprietes
	Importe dans chaque part/product la valeure des attributs documentés dans le fichier B_ExtractNom

D_ExpotOrdo
	Exporte vers un template Excel les données des attributs des parts et products de l'assemblage
	selon l'association des attributs/colonnes défini dans le fichier texte

Classes:
-------
c_LignomOrdo : Collection des lignes de nomenclatures a déverser dans le template Excel
c_Product : Collection des produits (composnats, parts et products)
c_LNomencl, c_Attribut : Classe de la bibliMacrosVBA


Format du fichier texte des attributs:
-------------------------------------
	NomEnv;xxxxxxx		--Nom de l'environnement de travail
	TemplateOrdo;NomOrdostd		-- nom du template Excel
	Attrib;x_designation;8		-- Premier attribut avec la colonne du template Excel associée
	Attrib;x_designation anglais;0
	Attrib;CODE;4
	Attrib;Marque;0





